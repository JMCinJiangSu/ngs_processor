"""
NGS 下机数据自动化处理工具 — 入口
"""

import os
import sys
import glob
import time
import warnings

import pandas as pd
import openpyxl

from config import PRODUCTS, ProductConfig, detect_product
from filters import (
    check_qc, to_num,
    process_snvindel,
    process_hot_somatic, process_somatic, process_discard_rescue,
    flag_hd_pass,
)
from writer import (
    write_qc_report, write_qc_failitem,
    write_snvindel_review,
    write_hot_somatic_review, write_somatic_review, write_discard_review,
    write_cnv, write_amplicon_pivot, write_passthrough, write_hd_pass,
)

warnings.filterwarnings("ignore")


def load_fake_pos(work_dir: str, cfg: ProductConfig) -> pd.DataFrame | None:
    candidates = glob.glob(
        os.path.join(work_dir, "**", cfg.fake_pos_filename), recursive=True
    )
    if not candidates:
        print(f"  ⚠ 未找到假阳文件 {cfg.fake_pos_filename}，假阳性检查将跳过")
        return None
    try:
        sheets = pd.read_excel(candidates[0], sheet_name=None, dtype=str)
        df = list(sheets.values())[0]
        print(f"  已加载假阳文件: {os.path.basename(candidates[0])}（{len(df)} 条）")
        return df
    except Exception as e:
        print(f"  ⚠ 假阳文件读取失败: {e}")
        return None


def process_file(ngs_path: str, fake_pos_df: pd.DataFrame | None, cfg: ProductConfig):
    t0 = time.time()
    print(f"\n{'='*60}")
    print(f"产品: {cfg.name}  文件: {os.path.basename(ngs_path)}")
    print(f"{'='*60}")

    def read_sheet(name):
        try:
            return pd.read_excel(ngs_path, sheet_name=name)
        except Exception:
            return pd.DataFrame()

    # ── 读取 ─────────────────────────────────────────────────
    t_read = time.time()
    summary_df = read_sheet("Summary")
    cnv_df     = read_sheet("CNV")      if cfg.has_cnv      else pd.DataFrame()
    amplicon_df= read_sheet("AmpliconStat") if cfg.has_amplicon else pd.DataFrame()
    hd_pass_df = read_sheet("HD_pass") if "HD_pass" in cfg.passthrough_sheets else pd.DataFrame()
    passthrough_data = {s: read_sheet(s) for s in cfg.passthrough_sheets if s != "HD_pass"}

    if cfg.has_snvindel:
        snv_df     = read_sheet("SNVIndel")
        discard_df = read_sheet("SNVIndelDiscard")
        hot_df = somatic_df = pd.DataFrame()
    elif cfg.has_snvindel_split:
        hot_df     = read_sheet("SNVIndelHotSomatic")
        somatic_df = read_sheet("SNVIndelSomatic")
        discard_df = read_sheet("SNVIndelDiscard")
        snv_df     = pd.DataFrame()
    else:
        snv_df = hot_df = somatic_df = discard_df = pd.DataFrame()

    print(f"  读取耗时: {time.time() - t_read:.1f}s")

    # ── QC ───────────────────────────────────────────────────
    qc_df, fail_dict, risk_dict = pd.DataFrame(), {}, {}
    if summary_df.empty:
        print("  ⚠ 未找到 Summary Sheet，跳过质控")
    else:
        for col in cfg.qc_rules:
            if col in summary_df.columns:
                summary_df[col] = summary_df[col].apply(
                    lambda x: to_num(x) if to_num(x) is not None else x
                )
        qc_df, fail_dict, risk_dict = check_qc(summary_df, cfg)
        print(f"  样本数: {len(summary_df)}  "
              f"合格: {len(summary_df)-len(fail_dict)-len(risk_dict)}  "
              f"风险: {len(risk_dict)}  不合格: {len(fail_dict)}")

    # ── SNVIndel 处理 ─────────────────────────────────────────
    if cfg.has_snvindel and not snv_df.empty:
        snv_df = process_snvindel(snv_df, discard_df, fake_pos_df, cfg)
        print(f"  SNVIndel 过滤后行数: {len(snv_df)}")

    if cfg.has_snvindel_split:
        if not hot_df.empty:
            hot_df = process_hot_somatic(hot_df, discard_df, fake_pos_df, cfg)
            print(f"  HotSomatic 过滤后行数: {len(hot_df)}")
        if not somatic_df.empty:
            somatic_df = process_somatic(somatic_df, discard_df, fake_pos_df, cfg)
            print(f"  Somatic 过滤后行数: {len(somatic_df)}")
        rescued_df = process_discard_rescue(discard_df, cfg)
        print(f"  Discard 二次筛选行数: {len(rescued_df)}")

    # ── 写出 ─────────────────────────────────────────────────
    t_write = time.time()
    base     = os.path.splitext(os.path.basename(ngs_path))[0]
    out_path = os.path.join(os.path.dirname(ngs_path), f"{base}_Report.xlsx")

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    if not qc_df.empty:
        write_qc_report(out_wb, qc_df, fail_dict, risk_dict, cfg)
    write_qc_failitem(out_wb, fail_dict, risk_dict)

    if cfg.has_snvindel and not snv_df.empty:
        write_snvindel_review(out_wb, snv_df, cfg)

    if cfg.has_snvindel_split:
        write_hot_somatic_review(out_wb, hot_df, cfg)
        write_somatic_review(out_wb, somatic_df, cfg)
        write_discard_review(out_wb, rescued_df, cfg)

    if cfg.has_cnv:
        write_cnv(out_wb, cnv_df)

    for sheet_name, df in passthrough_data.items():
        if not df.empty:
            write_passthrough(out_wb, sheet_name, df)
            print(f"  已输出 Sheet: {sheet_name}")

    if not hd_pass_df.empty:
        hd_pass_df = flag_hd_pass(hd_pass_df, summary_df)
        write_hd_pass(out_wb, hd_pass_df)
        hd_yes = (hd_pass_df["HD_Flag"] == "Yes").sum()
        print(f"  HD_pass: 共 {len(hd_pass_df)} 行，HD_Flag=Yes: {hd_yes} 行")

    if cfg.has_amplicon:
        write_amplicon_pivot(out_wb, amplicon_df)

    out_wb.save(out_path)
    print(f"  写出耗时: {time.time() - t_write:.1f}s")
    print(f"  ✔ 报告已生成: {os.path.basename(out_path)}  （总耗时 {time.time()-t0:.1f}s）")

    if fail_dict or risk_dict:
        print("\n  【质控提示】")
        for sample, fails in fail_dict.items():
            print(f"    [不合格] {sample} {', '.join(fails)}")
        for sample, risks in risk_dict.items():
            if sample not in fail_dict:
                print(f"    [风险]   {sample} {', '.join(risks)}")


def main():
    work_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.getcwd()
    print("NGS 下机数据自动化处理工具")
    print(f"工作目录: {work_dir}")
    print(f"已注册产品: {[p.name for p in PRODUCTS]}")

    all_xlsx = glob.glob(os.path.join(work_dir, "**", "*.xlsx"), recursive=True)
    all_xlsx = [f for f in all_xlsx if "_Report" not in os.path.basename(f)]

    product_files: dict[str, list[str]] = {}
    unmatched = []
    for path in all_xlsx:
        cfg = detect_product(os.path.basename(path))
        if cfg:
            product_files.setdefault(cfg.name, []).append(path)
        else:
            unmatched.append(path)

    if unmatched:
        print(f"\n⚠ 以下文件未匹配到任何产品，跳过：")
        for f in unmatched:
            print(f"  {os.path.basename(f)}")

    total_files = sum(len(v) for v in product_files.values())
    if total_files == 0:
        print("\n⚠ 未找到任何可处理的下机数据文件")
        input("\n按回车键退出...")
        return

    print(f"\n共找到 {total_files} 个下机数据文件")

    fake_pos_cache: dict[str, pd.DataFrame | None] = {}
    for cfg in PRODUCTS:
        if cfg.name in product_files:
            fake_pos_cache[cfg.name] = load_fake_pos(work_dir, cfg)

    total_t0 = time.time()
    file_times = []
    for cfg in PRODUCTS:
        for path in product_files.get(cfg.name, []):
            ft0 = time.time()
            try:
                process_file(path, fake_pos_cache[cfg.name], cfg)
            except Exception as e:
                print(f"\n  ✗ 处理 {os.path.basename(path)} 时出错: {e}")
                import traceback; traceback.print_exc()
            file_times.append((os.path.basename(path), time.time() - ft0))

    print("\n" + "=" * 60)
    if len(file_times) > 1:
        print("各文件耗时：")
        for fname, t in file_times:
            print(f"  {fname}: {t:.1f}s")
    print(f"全部处理完成！总耗时 {time.time()-total_t0:.1f}s")
    input("按回车键退出...")


if __name__ == "__main__":
    main()
