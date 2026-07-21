"""
产品配置文件
新增产品只需在 PRODUCTS 列表中添加一个 ProductConfig 对象，无需修改其他文件。
"""

from dataclasses import dataclass, field

# ══════════════════════════════════════════════════════════════
# QC 规则类型
#
# 单阈值（原有）：{列名: (">=", 合格值)}
# 双阈值（新增）：{列名: (">=", 合格值, 风险值)}
#   val >= 合格值 → 绿（合格）
#   合格值 > val >= 风险值 → 黄（风险）
#   val < 风险值 → 红（不合格）
# ══════════════════════════════════════════════════════════════

# 服务器映射表（用于生成bam文件路径）
server_dict = {
    "HW2B4N2": "server-2-1", # 上海肺科
    "1G27M74": "server-5-142", # 上海十院
    "JW1DWF3": "server-4-84", # 华东医院
    "FGN7X34": "server-5-121", # 复旦中山
    "G5KL6S3": "server-6-3", # 青岛附属医院
}

@dataclass
class ProductConfig:
    # ── 基本信息 ──────────────────────────────────────────────
    name: str                   # 产品名称（用于日志显示）
    file_keyword: str           # 下机文件名识别关键词
    fake_pos_filename: str      # 假阳文件名（放在同一工作目录）
    qc_sheet_name: str          # QC 统计表名称（下机表格中）
    cnv_sheet_name: str         # CNV 检测结果表名称（下机表格中）
    hd_sheet_name: str          # HD检测结果表名称（下机表格中）
    server_dict: dict = field(default_factory=lambda: server_dict.copy())  # 服务器映射表（用于生成bam文件路径）

   # ── QC 规则 ────────────────────────────────────────────────
    # 单阈值: {col: (">=", pass_val)}
    # 双阈值: {col: (">=", pass_val, risk_val)}
    qc_rules: dict = field(default_factory=dict)

    # ── 各 Sheet 保留列 ────────────────────────────────────────
    qc_report_cols: list = field(default_factory=list)
    snv_review_cols: list = field(default_factory=list)
    fusion_cols: list = field(default_factory=list)

    # ── SNVIndel 过滤关键词（Tags 列） ─────────────────────────
    tag_filters: list = field(default_factory=lambda: ["Black_list", "Polymorphism"])

    # ── AltDepth 低深度阈值 ────────────────────────────────────
    low_altdepth_threshold: int = 30

    # ── SNVIndel 复核标记规则 ─────────────────────────
    review_freq_threshold_polymer_str: float = 0.3 # Tags含Polymer或STR
    review_freq_threshold_other: float = 0.4        # Tags不含Polymer或STR
    review_significance_values: list = field(default_factory=list)  # 复核标记的Significance值列表
    review_tag_keywords: list = field(default_factory=list)


    # ── 直接透传原始数据的 sheet 列表 ─────────────────────────
    passthrough_sheets: list = field(default_factory=list)

    # ── 功能开关 ───────────────────────────────────────────────
    has_snvindel: bool = True          # 标准 SNVIndel 单表模式
    has_snvindel_split: bool = False   # 新模式：HotSomatic + Somatic + Discard 三表
    has_amplicon: bool = True

    # ── 新模式：HotSomatic Freq 低频标记阈值 ───────────────────
    hot_somatic_low_freq: float = 0.01
    var_us_threshold: int = 20

    # ── 新模式：SNVIndelDiscard 二次筛选规则 ───────────────────
    # AND 条件：Significance in significance_rescue_values
    #           AND Freq >= freq_rescue_min
    #           AND (Tags 含 rescue_tag_keywords 中任一 OR Gene 在 rescue_genes 中)
    significance_rescue_values: list = field(default_factory=lambda: ["3", "4", "5"])
    discard_rescue_sig_gene: list = field(default_factory=lambda: ["4", "5"])
    freq_rescue_min: float = 0.005
    rescue_tag_keywords: list = field(default_factory=list)
    rescue_genes: list = field(default_factory=list)   # 指定基因列表
    discard_tag_exclude: list = field(default_factory=list) # 排除含指定字段的行
    
    # SNVIndelGermNonIC 筛选规则
    # 
    snv_type: list = field(default_factory=lambda: ["Nonsense_Mutation", "FrameShift_Duplication", "Splicing", "FrameShift_Deletion"]) 

    # ── 新模式：各 sheet 保留列（可与 snv_review_cols 不同） ───
    hot_somatic_review_cols: list = field(default_factory=list)
    somatic_review_cols: list = field(default_factory=list)
    discard_review_cols: list = field(default_factory=list)


# ══════════════════════════════════════════════════════════════
# 产品注册表（按此顺序匹配文件名，先匹配先生效）
# ══════════════════════════════════════════════════════════════

PRODUCTS: list[ProductConfig] = [

    # OncoPro组织
    ProductConfig(
        name="OncoPro",
        file_keyword="ADXHS-OncoPro",
        fake_pos_filename="CP200_db.xlsx",
        qc_sheet_name="Summary",
        cnv_sheet_name="CNV",
        hd_sheet_name="HD_pass",
        qc_rules={
            "CleanQ30":          (">=", 0.75),
            "Depth_CDS":         (">=", 400.0),
            "RNA-Control":       (">=", 10.0),
            "Coverage(50x)_SNP": (">=", 0.90),
        },
        qc_report_cols=[
            "Sample", "CleanData", "CleanQ30", "Depth_CDS",
            "Coverage(50x)_SNP", "RNA-Control",
            "MSI_Ratio", "MSI_Num", "MSI_State",
            "ContaRatio", "ContaStat",
        ],
        snv_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tags", "Depth", "Freq", "AltDepth",
            "Gene", "Type", "CDSChange", "Amplicon", "Plus", "Minus", "Significance",
        ],
        fusion_cols=[
            "Sample", "Fusion", "Copies", "Tags",
            "5'Chr", "5'BreakPoint", "3'Chr", "3'BreakPoint",
        ],
        discard_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tags", "Depth", "Freq", "AltDepth",
            "Gene", "Type", "CDSChange", "Amplicon", "Plus", "Minus", "Significance",
        ],
        tag_filters=["Black_list", "Polymorphism"],
        low_altdepth_threshold=30,
        # Discard 捞回
        significance_rescue_values=[4, 5],
        freq_rescue_min=0.0001,
        rescue_tag_keywords=["HotSpot"],
        rescue_genes=[
            "ALK", "BRAF", "EGFR", "ERBB2", "ROS1",
            "KRAS", "MET", "NRAS", "PIK3CA", "RET"
            # 根据实际需求补充
        ],
        passthrough_sheets=["Fusion"],
        has_snvindel=True,
        has_snvindel_split=False,
        has_amplicon=True,
    ),
    # Classic Panel
    ProductConfig(
        name="Classic Panel",
        file_keyword="ADXHS-Classic",
        fake_pos_filename="",
        qc_sheet_name="Summary",
        cnv_sheet_name="CNA",
        hd_sheet_name="",
        qc_rules={
            "CleanQ30":          (">=", 0.75),
            "Depth":         (">=", 400.0),
            "RNA-Control":       (">=", 20.0),
        },
        qc_report_cols=[
            "Sample", "CleanData", "CleanQ30", "Depth", "RNA-Control",
            "MSI_Ratio", "MSI_Num", "MSI_State", "mRNA-imbalance"
        ],
        snv_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tags", "Depth", "Freq", "AltDepth",
            "Gene", "Type", "CDSChange", "Amplicon", "Plus", "Minus", "ClinVar_CLNSIG",
        ],
        fusion_cols=[
            "Sample", "Fusion", "Copies",
            "5'Chr", "5'BreakPoint", "3'Chr", "3'BreakPoint",
        ],
        tag_filters=["Black_list", "Polymorphism"],
        low_altdepth_threshold=30,
        passthrough_sheets=["NEWFusion", "Fusion"],
        significance_rescue_values=[4, 5],
        freq_rescue_min=0.0001,
        rescue_tag_keywords=["HotSpot"],
        has_snvindel=True,
        has_snvindel_split=False,
        has_amplicon=True,
    ),
    # Master DNA
    ProductConfig(
        name="Master Panel DNA",
        file_keyword="ADXMaster-DNA",
        fake_pos_filename="MP_db.xlsx",
        qc_sheet_name="QC",
        cnv_sheet_name="CNV",
        hd_sheet_name="",
        qc_rules={
            "CleanQ30":          (">=", 0.75),
            "Coverage":         (">=", 0.95),
            "HotUNIQUni-20%":       (">=", 0.9, 0.8),
            "NonHotUNIQUni-20%":       (">=", 0.8, 0.7),
            "HotUNIQDepth":       (">=", 800.0, 600.0),
            "NonHotUNIQDepth":       (">=", 400.0, 300.0),
            "MSIScore":           ("risk_between", 250.0, 150.0)
        },
        qc_report_cols=[
            "Sample", "cleanData", "cleanQ30", "Coverage", "Mapping", "HotUNIQUni-20%", "NonHotUNIQUni-20%", "HotUNIQDepth", "NonHotUNIQDepth",
            "MSIScore", "MSIStat", "TMB", "SSBCDepth", "ContaRatio", "ContaStat"
        ],
        snv_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tag", "Depth", "Freq", "Significance",
            "Gene", "Type", "CDSChange", "Depth_US", "Freq_US", "Var_US", "Depth_SS", "Freq_SS", "Var_SS", "Depth_DS", "Freq_DS", "Var_DS"
        ],
        fusion_cols=[
            "Sample", "Fusion", "Copies",
            "5'Chr", "5'BreakPoint", "3'Chr", "3'BreakPoint",
        ],
        tag_filters=[],
        low_altdepth_threshold=30,
        passthrough_sheets=["HD", "Fusion"],
        has_snvindel=False,
        has_snvindel_split=True,
        has_amplicon=False,
        # HotSomatic Freq 低频标记
        hot_somatic_low_freq=0.01,
        var_us_threshold=20,
        hot_somatic_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tag", "Depth", "Freq", "Significance",
            "Gene", "Type", "CDSChange", "Depth_US", "Freq_US", "Var_US", "Depth_SS", "Freq_SS", "Var_SS", "Depth_DS", "Freq_DS", "Var_DS"
        ],
        somatic_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tag", "Depth", "Freq", "Significance",
            "Gene", "Type", "CDSChange", "Depth_US", "Freq_US", "Var_US", "Depth_SS", "Freq_SS", "Var_SS", "Depth_DS", "Freq_DS", "Var_DS"
        ],
        discard_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tag", "Depth", "Freq", "Significance",
            "Gene", "Type", "CDSChange", "Depth_US", "Freq_US", "Var_US", "Depth_SS", "Freq_SS", "Var_SS", "Depth_DS", "Freq_DS", "Var_DS"
        ],
        # Discard 二次筛选（三条件 AND）
        significance_rescue_values=["3", "4", "5"],
        discard_rescue_sig_gene=["4", "5"],
        freq_rescue_min=0.005,
        rescue_tag_keywords=["HotSpot"],
        rescue_genes=[
            "ALK", "BRAF", "CTNNB1", "EGFR", "ERBB2", "FGFR1",
            "FGFR2", "FGFR3", "FGFR4", "IDH1", "IDH2", "KIT",
            "KRAS", "MET", "NRAS", "NTRK1", "NTRK2", "NTRK3",
            "PDGFRA", "PIK3CA", "POLE", "PTEN", "RB1", "RET",
            "ROS1", "TP53"
            # 根据实际需求补充
        ],
    ),
    # Master RNA
    ProductConfig(
        name="Master Panel RNA",
        file_keyword="ADXMaster-RNA",
        fake_pos_filename="",
        qc_sheet_name="DataProduction",
        cnv_sheet_name="",
        hd_sheet_name="",
        qc_rules={
            "CleanQ30":          (">=", 0.75),
            "Mapping":         (">=", 0.80, 0.70),
            "End2SenseRate":       (">=", 0.9, 0.80),
            "effectiveReads":       (">=", 3000000),
            "controlReads":       (">=", 2000.0),
        },
        qc_report_cols=[
            "Sample", "CleanData", "CleanQ30", "effectiveReads", "Mapping", "iSize", "End2SenseRate"
        ],
        snv_review_cols=[],
        fusion_cols=[],
        tag_filters=[],
        low_altdepth_threshold=30,
        passthrough_sheets=["FusionDiscard", "Fusion", "Expression", "EBV", "GEP", "TME", "CUP"],
        has_snvindel=False,
        has_snvindel_split=False,
        has_amplicon=False,
    ),
    # BPTM Plus 组织
    ProductConfig(
        name="BPTM Plus 组织",
        file_keyword="ADXHS-tBPTMplus",
        fake_pos_filename="",
        qc_sheet_name="Summary",
        cnv_sheet_name="",
        hd_sheet_name="",
        qc_rules={
            "CleanQ30":          (">=", 0.75),
            "Depth":         (">=", 300.0),
            "MSINum":           ("risk_between", 0.24, 0.12)
        },
        qc_report_cols=[
            "Sample", "CleanData", "CleanQ30", "Depth",
            "MappingRate", "RawDepth", "Uniformity(20%)",
            "MSI_Ratio", "MSI_Num", "MSI_State",
        ],
        snv_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tags", "Depth", "Freq", "AltDepth",
            "Gene", "Type", "CDSChange", "Amplicon", "Plus", "Minus", "Significance",
        ],
        discard_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tags", "Depth", "Freq", "AltDepth",
            "Gene", "Type", "CDSChange", "Amplicon", "Plus", "Minus", "Significance",
        ],
        tag_filters=["Black_list", "Polymorphism"],
        low_altdepth_threshold=30,
        # Discard 捞回
        significance_rescue_values=[4, 5],
        freq_rescue_min=0.0015,
        rescue_tag_keywords=[],
        discard_tag_exclude=["Bad", "Polymorphism", "OutOfReg"],
        rescue_genes=[
            "POLE", "BRCA1", "BRCA2", "MLH1", "MSH2",
            "MSH6", "PMS2", "TP53"
            # 根据实际需求补充
        ],
        passthrough_sheets=[],
        has_snvindel=True,
        has_snvindel_split=False,
        has_amplicon=False,
    ),
    # 遗传150
    ProductConfig(
        name="遗传150",
        file_keyword="ADXHS-gHC",
        fake_pos_filename="gHC_db.xlsx",
        qc_sheet_name="Summary",
        cnv_sheet_name="CNV",
        hd_sheet_name="",
        qc_rules={
            "CleanQ30":          (">=", 0.75),
            "Coverage(20X)":         (">=", 0.99),
        },
        qc_report_cols=[
            "Sample", "CleanData", "CleanQ30", "Coverage(20X)", "Depth",
            "MappingRate", "RawDepth", "Uniformity(20%)"
        ],
        snv_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tags", "Depth", "Freq", "AltDepth",
            "Gene", "Type", "CDSChange", "Amplicon", "Plus", "Minus", "Significance",
        ],
        discard_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tags", "Depth", "Freq", "AltDepth",
            "Gene", "Type", "CDSChange", "Amplicon", "Plus", "Minus", "Significance",
        ],
        tag_filters=["Black_list", "Polymorphism"],
        low_altdepth_threshold=30,
        review_freq_threshold_polymer_str=0.3,
        review_freq_threshold_other=0.4,
        review_significance_values=[3, 4, 5],
        review_tag_keywords=["Polymer", "STR"],
        # Discard 捞回
        significance_rescue_values=[4, 5],
        freq_rescue_min=0.0015,
        rescue_tag_keywords=[],
        discard_tag_exclude=[],
        rescue_genes=[],
        passthrough_sheets=["Checklist"],
        has_snvindel=True,
        has_snvindel_split=False,
        has_amplicon=False,
    ),
    # BRCA
    ProductConfig(
        name="tBRCA",
        file_keyword="BRCA",
        fake_pos_filename="",
        qc_sheet_name="Summary",
        cnv_sheet_name="",
        hd_sheet_name="",
        qc_rules={
            "CleanQ30":          (">=", 0.75),
            "MappingRate":         (">=", 0.85),
            "Coverage":         ("==", 1),
            "Depth":         (">=", 300),
            "MinDepth":         (">=", 50),
        },
        qc_report_cols=[
            "Sample", "CleanData", "CleanQ30", "Coverage", "MinDepth", "Depth",
            "MappingRate", "RawDepth", "Uniformity(20%)"
        ],
        snv_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tags", "Depth", "Freq", "AltDepth",
            "Gene", "Type", "CDSChange", "Amplicon", "Plus", "Minus", "Significance",
        ],
        discard_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tags", "Depth", "Freq", "AltDepth",
            "Gene", "Type", "CDSChange", "Amplicon", "Plus", "Minus", "Significance",
        ],
        tag_filters=["Black_list", "Polymorphism"],
        low_altdepth_threshold=30,
        review_freq_threshold_polymer_str=0.35,
        review_tag_keywords=["Polymer", "STR"],
        # Discard 捞回
        significance_rescue_values=["4", "5"],
        freq_rescue_min=0.005,
        rescue_tag_keywords=[],
        discard_tag_exclude=["Bad", "Polymorphism", "OutOfReg"],
        rescue_genes=[],
        passthrough_sheets=[],
        has_snvindel=True,
        has_snvindel_split=False,
        has_amplicon=False,
    ),
    # HRD
    ProductConfig(
        name="HRD",
        file_keyword="ADXHS-tHRD",
        fake_pos_filename="",
        qc_sheet_name="Summary",
        cnv_sheet_name="",
        hd_sheet_name="HD",
        qc_rules={
            "CleanQ30":          (">=", 0.75),
            "Coverage(180x)_BRCA":         (">=", 0.95),
            "Coverage(180x)_CDS":         (">=", 0.95),
            "BAFNoise":         ("<=", 0.05),
            "DepthNoise":         ("<=", 0.35),
        },
        qc_report_cols=[
            "Sample", "CleanData", "CleanQ30", "Coverage(180x)_BRCA", "Coverage(180x)_CDS", "BAFNoise", "DepthNoise",
            "MappingRate", "GSS", "ContaRatio", "ContaStat"
        ],
        snv_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tags", "Depth", "Freq", "AltDepth",
            "Gene", "Type", "CDSChange", "Amplicon", "Plus", "Minus", "Significance",
        ],
        discard_review_cols=[
            "Sample", "Chr", "Start", "End", "Ref", "Alt",
            "Tags", "Depth", "Freq", "AltDepth",
            "Gene", "Type", "CDSChange", "Amplicon", "Plus", "Minus", "Significance",
        ],
        tag_filters=["Black_list", "Polymorphism"],
        low_altdepth_threshold=30,
        # Discard 捞回
        significance_rescue_values=["4", "5"],
        freq_rescue_min=0.015,
        rescue_tag_keywords=[],
        discard_tag_exclude=["Bad", "Polymorphism", "OutOfReg"],
        rescue_genes=[],
        passthrough_sheets=[],
        has_snvindel=True,
        has_snvindel_split=False,
        has_amplicon=False,
    ),
    # BRCA V1
    ProductConfig(
        name="BRCA V1",
        file_keyword="_raw",
        fake_pos_filename="BRCA_SNP.xlsx",
        qc_sheet_name="Summary",
        cnv_sheet_name="",
        hd_sheet_name="",
        qc_rules={},
        qc_report_cols=[],
        snv_review_cols=["Gene", "Tags", "Chr:Start-End", "Ref", "Var", "Context", "Depth", "Frequency", "Ploidy",
                         "Type", "CDSChange", "dbID", "Significance", "Source", "Reference", "Amplicon"],
        discard_review_cols=[],
        tag_filters=[],
        low_altdepth_threshold=30,
        # Discard 捞回
        significance_rescue_values=[4, 5],
        freq_rescue_min=0.015,
        rescue_tag_keywords=[],
        discard_tag_exclude=[],
        rescue_genes=[],
        passthrough_sheets=["Summary", "AmpliconStat"],
        has_snvindel=False,
        has_snvindel_split=False,
        has_amplicon=False,
    ),
]

def detect_product(filename: str) -> ProductConfig | None:
    """根据文件名识别产品，返回对应配置；无匹配则返回 None。"""
    for product in PRODUCTS:
        if product.file_keyword in filename:
            return product
    return None
