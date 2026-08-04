# NGS 下机数据自动化处理工具使用手册

## 1. 工具简介

本项目用于批量处理 NGS 下机 Excel 文件，自动完成产品识别、质控检查、SNV/Indel 结果复核标记、Discard 二次筛选、CNV/HD 标记、Amplicon 透视表生成以及指定 Sheet 透传，最终输出复核用 Excel 报告。

程序入口为 `main.py`，核心配置集中在 `config.py`，过滤/标记逻辑在 `filters.py`，Excel 写出逻辑在 `writer.py`。

## 2. 目录结构

建议运行目录保持如下结构：

```text
ngs_processor/
├── main.py              # 程序入口
├── config.py            # 产品配置、QC 阈值、输出列配置
├── filters.py           # 过滤、统计、标记逻辑
├── writer.py            # Excel 报告写出逻辑
├── main.spec            # PyInstaller 打包配置
├── excel/               # 存放待处理下机 Excel 文件
├── fake_db/             # 存放假阳性数据库 Excel 文件
└── result/              # 输出目录，程序生成报告前需确保存在
```

> 注意：程序会扫描运行目录下 `excel/*.xlsx` 文件，并跳过文件名包含 `_Report` 的文件。

## 3. 环境要求

### 3.1 Python 运行环境

推荐使用 Python 3.10 或更高版本。项目依赖以下第三方库：

- `pandas`
- `openpyxl`

安装依赖：

```bash
pip install pandas openpyxl
```

### 3.2 Windows 可执行文件运行

项目包含 `main.spec`，可使用 PyInstaller 打包为控制台程序：

```bash
pip install pyinstaller pandas openpyxl
pyinstaller main.spec
```

打包后将 `excel/`、`fake_db/`、`result/` 放在可执行文件同级目录下运行。

## 4. 输入文件准备

### 4.1 下机数据文件

将待处理的下机 Excel 文件放入：

```text
excel/
```

程序根据文件名关键字自动识别产品类型。当前支持产品如下：

| 产品名称 | 文件名识别关键字 | QC Sheet | CNV Sheet | HD Sheet | 假阳性数据库 |
| --- | --- | --- | --- | --- | --- |
| OncoPro | `ADXHS-OncoPro` | `Summary` | `CNV` | `HD_pass` | `CP200_db.xlsx` |
| Classic Panel | `ADXHS-Classic` | `Summary` | `CNA` | 无 | 无 |
| Master Panel DNA | `ADXMaster-DNA` | `QC` | `CNV` | 无 | `MP_db.xlsx` |
| Master Panel RNA | `ADXMaster-RNA` | `DataProduction` | 无 | 无 | 无 |
| BPTM Plus 组织 | `ADXHS-tBPTMplus` | `Summary` | 无 | 无 | 无 |
| 遗传150 | `ADXHS-gHC` | `Summary` | `CNV` | 无 | `gHC_db.xlsx` |
| tBRCA | `BRCA` | `Summary` | 无 | 无 | 无 |
| HRD | `ADXHS-tHRD` | `Summary` | 无 | `HD` | 无 |
| BRCA V1 | `_raw` | `Summary` | 无 | 无 | `BRCA_SNP.xlsx` |

未匹配到以上关键字的 Excel 文件会被跳过，并在控制台输出提示。

### 4.2 假阳性数据库文件

需要假阳性判断的产品，应将对应文件放入：

```text
fake_db/
```

假阳性库读取第一个 Sheet，并使用 `Chr:Start-End` 或 `CDSChange` 字段作为匹配键。若未找到对应假阳性文件，程序会继续运行，但跳过假阳性检查。

### 4.3 result 输出目录

程序输出报告到：

```text
result/
```

请在运行前确认该目录已存在，否则保存报告时可能报错。

## 5. 使用方法

### 5.1 命令行运行

在项目根目录执行：

```bash
python main.py
```

运行流程：

1. 打印工作目录。
2. 整理 `excel/` 目录：将子目录中的 `_raw.xlsx` 移到 `excel/` 根目录，并删除这些子目录。
3. 扫描 `excel/*.xlsx`。
4. 根据文件名识别产品。
5. 按产品加载假阳性数据库。
6. 逐个文件处理并输出报告。
7. 控制台显示总耗时，按回车退出。

### 5.2 可执行文件运行

若使用 PyInstaller 打包，双击或在命令行运行 `main.exe`。程序的工作目录为可执行文件所在目录。

## 6. 输出文件说明

### 6.1 输出文件命名

普通产品输出：

```text
result/<原始文件名>_Report.xlsx
```

BRCA V1 输出：

```text
result/<原始文件名>.xlsx
```

BRCA V1 仅在发现假阳性 SNP 位点时保存输出文件。

### 6.2 常见输出 Sheet

| Sheet 名称 | 说明 |
| --- | --- |
| `QC_Report` | 样本 QC 汇总，包含颜色标记和 bam_path 字段 |
| `QC_FailItems` | QC 不合格/风险样本汇总；BRCA V1 不输出该 Sheet |
| `SNVIndel_Review` | 单表 SNVIndel 模式的变异复核表 |
| `HotSomatic_Review` | 三表模式下 HotSomatic 复核表 |
| `Somatic_Review` | 三表模式下 Somatic 复核表 |
| `GermNonIC_Review` | 三表模式下 GermNonIC 筛选复核表 |
| `Discard_Review` | SNVIndelDiscard 二次筛选结果 |
| `CNV` | CNV 结果，并对低置信/需复核 CNV 进行橙色标记 |
| `HD_pass` | HD 结果，并对需复核行进行橙色标记 |
| `AmpliconStat` | 按 Amplicon 和 Sample 生成 RoT 透视表 |
| 透传 Sheet | 配置中指定的原始 Sheet 直接写入结果文件 |

## 7. 主要处理规则

### 7.1 QC 检查

QC 阈值在 `config.py` 的 `PRODUCTS` 中按产品配置。支持两类规则：

- 单阈值：例如 `(">=", 0.75)`，大于等于阈值合格，否则不合格。
- 双阈值：例如 `(">=", 合格值, 风险值)`，达到合格值为合格，低于合格值但达到风险值为风险，低于风险值为不合格。
- 区间风险：`("risk_between", 高值, 低值)`，落在区间内标记为风险。

输出颜色：

- 绿色：合格。
- 黄色：风险。
- 红色：不合格。

### 7.2 SNV/Indel 过滤与标记

程序根据产品配置处理单表或三表模式：

- 单表模式：读取 `SNVIndel` 和 `SNVIndelDiscard`。
- 三表模式：读取 `SNVIndelHotSomatic`、`SNVIndelSomatic`、`SNVIndelDiscard`、`SNVIndelGermNonIC`。

常见处理内容：

- 根据 `Tags` 或 `Tag` 过滤黑名单、常见多态等配置关键词。
- 统计同一变异在 `SNVIndelDiscard` 中出现次数，生成 `DisCard_Count`。
- 统计同一变异在当前 SNV 表中出现次数，生成 `SNVIndel_Count`。
- 与假阳性数据库匹配，生成 `IsFakePositive`。
- 根据 `AltDepth` 低深度阈值生成 `LowAltDepth_Flag`。
- HotSomatic 根据低频或 `Var_US` 阈值生成 `LowFreq_Flag`。

### 7.3 Discard 二次筛选

程序会从 `SNVIndelDiscard` 中筛选需要人工复核的变异。筛选逻辑由产品配置控制，主要依据：

- `Significance` 等级。
- `Freq` 最低阈值。
- `Tag/Tags` 是否包含热点关键词。
- `Gene` 是否在指定抢救基因列表中。
- 是否命中排除关键词。

筛选结果写入 `Discard_Review`。

### 7.4 CNV 标记

CNV 输出中满足以下任一条件的行会被橙色标记：

- `Confidence` 为 `Low`。
- 程序生成的 `CNV_Flag` 为 `Yes`。

其中 `CNV_Flag` 的当前逻辑为：`CopyNum >= 10` 且 `Auto == "F"`。

### 7.5 HD 标记

HD Sheet 会根据程序逻辑生成 `HD_Flag`，`HD_Flag == "Yes"` 的行会在输出中以橙色标记，提示需要复核。

### 7.6 bam_path 生成

`QC_Report` 会追加 `bam_path` 字段。程序通过文件路径中的测序服务器关键字匹配服务器地址，并结合 `Summary/QC` 表中的 `Sample`、`Library`、`FlowCell_Lane` 生成 bam 文件链接。

当前服务器关键字映射在 `config.py` 的 `server_dict` 中维护。

## 8. 产品配置维护

新增或调整产品时，通常只需要修改 `config.py` 中的 `PRODUCTS` 列表，新增一个 `ProductConfig` 配置对象即可。

常用配置项：

| 配置项 | 说明 |
| --- | --- |
| `name` | 产品显示名称 |
| `file_keyword` | 文件名识别关键字 |
| `fake_pos_filename` | 假阳性数据库文件名 |
| `qc_sheet_name` | QC Sheet 名称 |
| `cnv_sheet_name` | CNV Sheet 名称 |
| `hd_sheet_name` | HD Sheet 名称 |
| `qc_rules` | QC 判断规则 |
| `qc_report_cols` | QC_Report 输出列 |
| `snv_review_cols` | SNVIndel_Review 输出列 |
| `tag_filters` | SNV 表过滤关键词 |
| `low_altdepth_threshold` | 低 AltDepth 标记阈值 |
| `passthrough_sheets` | 直接透传输出的 Sheet 名称 |
| `has_snvindel` | 是否使用单表 SNVIndel 模式 |
| `has_snvindel_split` | 是否使用 HotSomatic/Somatic/Discard 三表模式 |
| `has_amplicon` | 是否输出 AmpliconStat 透视表 |

## 9. 常见问题

### 9.1 提示“未找到任何可处理的下机数据文件”

请检查：

- Excel 文件是否放在 `excel/` 目录下。
- 文件扩展名是否为 `.xlsx`。
- 文件名是否包含已配置产品的识别关键字。
- 文件名是否包含 `_Report`，包含该字符串会被程序跳过。

### 9.2 提示“未找到假阳文件”

请检查对应假阳性数据库是否放在 `fake_db/` 目录下，且文件名与产品配置中的 `fake_pos_filename` 完全一致。该提示不会中断流程，只会跳过假阳性检查。

### 9.3 输出文件无法保存

请检查：

- `result/` 目录是否存在。
- 目标报告文件是否已被 Excel 打开。
- 当前用户是否有写入权限。

### 9.4 某个 Sheet 没有输出

通常是因为原始下机文件中不存在对应 Sheet，或该 Sheet 读取后为空。程序会跳过空数据 Sheet。

## 10. 注意事项

- 运行前建议关闭已打开的输出报告，避免文件占用导致保存失败。
- 文件名识别按 `PRODUCTS` 列表顺序匹配，若多个关键字可能同时命中，应将更精确的配置放在前面。
- `BRCA V1` 使用 `_raw` 作为识别关键字，需注意不要让无关文件名误命中。
- 程序会移动 `excel/` 子目录中的 `_raw.xlsx` 到 `excel/` 根目录，并删除这些子目录，请确认目录中没有需要保留的其他文件。
