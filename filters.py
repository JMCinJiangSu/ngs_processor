"""
过滤与统计逻辑
所有函数通过参数接收配置，不引用任何全局常量。
"""

import pandas as pd
from config import ProductConfig


def to_num(val):
    """安全转换为浮点数，去除 % 符号。"""
    try:
        return float(str(val).replace("%", "").strip())
    except Exception:
        return None


def normalize_cds(raw: str) -> str:
    """CDSChange规范化，去除空格，主要用于匹配和统计。"""
    return str(raw).strip()


def cds_key(df: pd.DataFrame) -> pd.Series:
    """所有统计的统一识别键：normalize_cds(CDSChange)。"""
    return df["CDSChange"].astype(str).map(normalize_cds)


# ══════════════════════════════════════════════════════════════
# QC 检查（支持单/双阈值）
# ══════════════════════════════════════════════════════════════

# check_qc 返回的样本状态
QC_PASS = "pass"
QC_RISK = "risk"
QC_FAIL = "fail"


def _eval_qc_col(val, rule: tuple) -> str:
    """
    单列 QC 判断。
    rule = (">=", pass_val)           → 单阈值
    rule = (">=", pass_val, risk_val) → 双阈值
    返回 QC_PASS / QC_RISK / QC_FAIL。
    """
    op = rule[0]
    
    if op == ">=":
        pass_val = rule[1]
        risk_val = rule[2] if len(rule) >= 3 else None
        if val >= pass_val:
            return QC_PASS
        if risk_val is not None and val >= risk_val:
            return QC_RISK
        return QC_FAIL
    elif op == "risk_between":
        hi, lo = rule[1], rule[2]
        return QC_RISK if lo <= val <= hi else QC_PASS
    return QC_PASS   # 未知操作符默认通过

def check_qc(summary_df: pd.DataFrame, cfg: ProductConfig):
    """
    按 cfg.qc_rules 检查每个样本，支持单/双阈值。

    返回:
      qc_df     — 仅保留 cfg.qc_report_cols 的 DataFrame
      fail_dict — {sample: [不合格描述]}   (QC_FAIL)
      risk_dict — {sample: [风险描述]}     (QC_RISK，双阈值专用)
    """
    available_cols = [c for c in cfg.qc_report_cols if c in summary_df.columns]
    qc_df = summary_df[available_cols].copy()

    fail_dict: dict[str, list[str]] = {}
    risk_dict: dict[str, list[str]] = {}
    
    for _, row in summary_df.iterrows():
        sample = str(row.get("Sample", "Unknown"))
        fails: list[str] = []
        risks: list[str] = []

        for col, rule in cfg.qc_rules.items():
            if col not in row.index:
                continue
            val = to_num(row[col])
            if val is None:
                fails.append(f"{col}(数据缺失)")
                continue

            status = _eval_qc_col(val, rule)
            pass_val = rule[1]
            risk_val = rule[2] if len(rule) >= 3 else None

            if status == QC_FAIL:
                ref = risk_val if risk_val is not None else pass_val
                fails.append(f"{col}={val}(不合格，标准≥{ref})")
            elif status == QC_RISK:
                fails_msg = f"{col}={val}(风险，建议≥{pass_val})"
                risks.append(fails_msg)

        if fails:
            fail_dict[sample] = fails
        if risks:
            risk_dict[sample] = risks

    return qc_df, fail_dict, risk_dict


# ══════════════════════════════════════════════════════════════
# SNVIndel 处理（标准单表模式，OncoPro）
# ══════════════════════════════════════════════════════════════

def _add_stats(snv_df: pd.DataFrame, discard_df: pd.DataFrame,
               fake_pos_df: pd.DataFrame | None, cfg: ProductConfig) -> pd.DataFrame:
    """公共统计列：DisCard_Count / SNVIndel_Count / IsFakePositive / LowAltDepth_Flag。"""
    snv_df["_key"] = cds_key(snv_df)

    discard_count: dict[str, int] = {}
    if discard_df is not None and not discard_df.empty:
        discard_count = cds_key(discard_df).value_counts().to_dict()
    snv_df["DisCard_Count"] = snv_df["_key"].map(lambda k: discard_count.get(k, 0))

    inner_count = snv_df["_key"].value_counts().to_dict()
    snv_df["SNVIndel_Count"] = snv_df["_key"].map(lambda k: inner_count.get(k, 0))

    fake_keys: set[str] = set()
    if fake_pos_df is not None and not fake_pos_df.empty:
        fake_keys = set(cds_key(fake_pos_df))
    snv_df["IsFakePositive"] = snv_df["_key"].map(lambda k: "是" if k in fake_keys else "否")

    threshold = cfg.low_altdepth_threshold
    if "AltDepth" in snv_df.columns:
        snv_df["LowAltDepth_Flag"] = snv_df["AltDepth"].apply(
            lambda x: "★低AltDepth"
            if (pd.notna(x) and to_num(x) is not None and to_num(x) <= threshold)
            else ""
        )

    snv_df.drop(columns=["_key"], inplace=True, errors="ignore")
    return snv_df

def process_snvindel(
    snv_df: pd.DataFrame,
    discard_df: pd.DataFrame,
    fake_pos_df: pd.DataFrame | None,
    cfg: ProductConfig,
) -> pd.DataFrame:
    """标准单表模式（OncoPro）：过滤黑名单 → 添加统计列。"""
    if "Tags" in snv_df.columns and cfg.tag_filters:
        pattern = "|".join(cfg.tag_filters)
        mask = snv_df["Tags"].astype(str).str.contains(pattern, case=False, na=False)
        snv_df = snv_df[~mask].copy()

    return _add_stats(snv_df, discard_df, fake_pos_df, cfg)

# ══════════════════════════════════════════════════════════════
# SNVIndel 处理（三表分离模式）
# ══════════════════════════════════════════════════════════════

def process_hot_somatic(
    hot_df: pd.DataFrame,
    discard_df: pd.DataFrame,
    fake_pos_df: pd.DataFrame | None,
    cfg: ProductConfig,
) -> pd.DataFrame:
    """
    HotSomatic 处理：
    - 过滤黑名单
    - 添加统计列
    - 添加 LowFreq_Flag（Freq < hot_somatic_low_freq）
    """
    if hot_df.empty:
        return hot_df

    if "Tag" in hot_df.columns and cfg.tag_filters:
        pattern = "|".join(cfg.tag_filters)
        mask = hot_df["Tag"].astype(str).str.contains(pattern, case=False, na=False)
        hot_df = hot_df[~mask].copy()

    hot_df = _add_stats(hot_df, discard_df, fake_pos_df, cfg)

    freq_threshold = cfg.hot_somatic_low_freq
    if "Freq" in hot_df.columns:
        hot_df["LowFreq_Flag"] = hot_df["Freq"].apply(
            lambda x: "★低频" if (pd.notna(x) and to_num(x) is not None and to_num(x) < freq_threshold)
            else ""
        )
    
    var_us_threshold = cfg.var_us_threshold
    if "Var_US" in hot_df.columns:
        hot_df["LowFreq_Flag"] = hot_df["Var_US"].apply(
            lambda x: "★低频" if (pd.notna(x) and to_num(x) is not None and to_num(x) < var_us_threshold)
            else ""
        )

    return hot_df

def process_somatic(
    somatic_df: pd.DataFrame,
    discard_df: pd.DataFrame,
    fake_pos_df: pd.DataFrame | None,
    cfg: ProductConfig,
) -> pd.DataFrame:
    """Somatic 处理：过滤黑名单 + 统计列（与标准模式相同）。"""
    if somatic_df.empty:
        return somatic_df

    if "Tag" in somatic_df.columns and cfg.tag_filters:
        pattern = "|".join(cfg.tag_filters)
        mask = somatic_df["Tag"].astype(str).str.contains(pattern, case=False, na=False)
        somatic_df = somatic_df[~mask].copy()

    return _add_stats(somatic_df, discard_df, fake_pos_df, cfg)


def process_discard_rescue(
    discard_df: pd.DataFrame,
    cfg: ProductConfig,
) -> pd.DataFrame:
    """
    从 SNVIndelDiscard 中二次筛选需要关注的变异（三条件 AND）：
      1. Significance 列值在 cfg.significance_rescue_values 中（4 或 5）
      2. Freq >= cfg.freq_rescue_min（0.005）
      3. Tags 含 cfg.rescue_tag_keywords 中任一关键词
         OR Gene 在 cfg.rescue_genes 中
    """
    if discard_df is None or discard_df.empty:
        return pd.DataFrame()

    df = discard_df.copy()

    # 条件1：Significance
    sig_vals = [str(v) for v in cfg.significance_rescue_values]
    if "Significance" in df.columns:
        cond1 = df["Significance"].astype(str).str.strip().isin(sig_vals)
    else:
        return pd.DataFrame()   # 缺少关键列，无法筛选

    # 条件2：Freq
    if "Freq" in df.columns:
        cond2 = df["Freq"].apply(
            lambda x: to_num(x) is not None and to_num(x) >= cfg.freq_rescue_min
        )
    else:
        return pd.DataFrame()

    # 条件3：Tag 含关键词 OR Gene 在指定列表
    tag_cond = pd.Series(False, index=df.index)
    if "Tag" in df.columns and cfg.rescue_tag_keywords:
        pattern = "|".join(cfg.rescue_tag_keywords)
        tag_cond = df["Tag"].astype(str).str.contains(pattern, case=False, na=False)

    gene_cond = pd.Series(False, index=df.index)
    if "Gene" in df.columns and cfg.rescue_genes:
        gene_cond = df["Gene"].astype(str).str.strip().isin(cfg.rescue_genes)

    cond3 = tag_cond | gene_cond

    rescued = df[cond1 & cond2 & cond3].copy()
    return rescued

def process_germnonic(
        germnonic_df: pd.DataFrame,
        cfg: ProductConfig
):
    if germnonic_df is None or germnonic_df.empty:
        return pd.DataFrame()
    
    df = germnonic_df.copy()

    # 条件1：Significance
    sig_vals = [str(v) for v in cfg.significance_rescue_values]
    if "Significance" in df.columns:
        cond1 = df["Significance"].astype(str).str.strip().isin(sig_vals)
    else:
        return pd.DataFrame()   # 缺少关键列，无法筛选
    # 条件2：Type
    if "Type" in df.columns and cfg.snv_type:
        cond2 = df["Type"].astype(str).str.strip().isin(cfg.snv_type)
    # 条件3：Tag
    tag_cond = pd.Series(False, index=df.index)
    if "Tag" in df.columns:
        tag_cond = df["Tag"].astype(str).str.contains("WBC", case=False, na=False)
    
    rescued = df[(cond1 | cond2) & ~tag_cond]
    return rescued


# ══════════════════════════════════════════════════════════════
# HD_pass 标记
# ══════════════════════════════════════════════════════════════

def flag_hd_pass(hd_df: pd.DataFrame, summary_df: pd.DataFrame) -> pd.DataFrame:
    """
    在 HD_pass 表中插入 HD_Flag 列。

    逻辑（按 Sample 关联 Summary 表）：
      1. Coverage(50x)_SNP < 0.9                          → HD_Flag = Yes
      2. Coverage(50x)_SNP >= 0.9:
           Cellularity >= 0.3
           AND RR <= 0.7                    (RR 在 HD_pass 表)
           AND (BAF >= 0.45 OR BAF == '-')  (BAF 在 HD_pass 表)
         满足以上全部                        → HD_Flag = No
         否则                               → HD_Flag = Yes
    """
    if hd_df is None or hd_df.empty:
        return hd_df

    hd_df = hd_df.copy()

    # 从 Summary 提取 Sample → Coverage(50x)_SNP、Cellularity 的映射
    cov_map: dict[str, float | None] = {}
    cell_map: dict[str, float | None] = {}
    if summary_df is not None and not summary_df.empty:
        for _, row in summary_df.iterrows():
            s = str(row.get("Sample", "")).strip()
            cov_map[s]  = to_num(row.get("Coverage(50x)_SNP"))
            cell_map[s] = to_num(row.get("Cellularity"))

    flags = []
    for _, row in hd_df.iterrows():
        sample = str(row.get("Sample", "")).strip()
        cov    = cov_map.get(sample)
        cell   = cell_map.get(sample)
        rr     = to_num(row.get("RR"))
        baf_raw = str(row.get("BAF", "")).strip()
        baf    = to_num(baf_raw)

        # 条件1：Coverage 不足，直接 Yes
        if cov is None or cov < 0.9:
            flags.append("Yes")
            continue

        # 条件2：Coverage 达标，细判
        cond_cell = (cell is not None and cell >= 0.3)
        cond_rr   = (rr   is not None and rr   <= 0.7)
        cond_baf  = (baf_raw == "-") or (baf is not None and baf >= 0.45)

        if cond_cell and cond_rr and cond_baf:
            flags.append("No")
        else:
            flags.append("Yes")

    hd_df["HD_Flag"] = flags
    return hd_df

# ══════════════════════════════════════════════════════════════
# CNV 标记
# ══════════════════════════════════════════════════════════════
def flag_cnv(cnv_df: pd.DataFrame) -> pd.DataFrame:
    """
    在CNV表里插入CNV_Flag列
    逻辑按照：
    CopyNum>=10 AND Auto = F
    """

    if cnv_df is None or cnv_df.empty:
        return cnv_df

    cnv_df = cnv_df.copy()

    flags = []
    for _, row in cnv_df.iterrows():
        copynum = to_num(row.get("CopyNum"))
        auto = row.get("Auto")

        if copynum >= 10 and auto == "F":
            flags.append("Yes")
        else:
            flags.append("No")
    cnv_df["CNV_Flag"] = flags

    return cnv_df