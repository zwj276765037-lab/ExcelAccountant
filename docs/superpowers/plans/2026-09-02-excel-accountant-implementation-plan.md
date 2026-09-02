# ExcelAccountant 实施计划

日期：2026-09-02

依据：`docs/superpowers/specs/2026-09-02-excel-accountant-design.md`

分支：`feature/local-accountant-app`

## 目标

交付一个完全离线的 Windows 桌面程序：读取普通 `.xlsx` 的单列完整十进制金额，为多个目标寻找源单元格互不重复的精确或近似分配；方案经用户勾选后输出带颜色和审计结果表的副本，近似输出必须显著记录差额。

## 任务 1：项目骨架与可重复环境

创建：

- `pyproject.toml`
- `requirements.txt`
- `requirements-dev.txt`
- `src/excel_accountant/`
- `tests/`
- `scripts/build.ps1`
- `README.md`

依赖：Python 3.12、PySide6、OR-Tools、openpyxl、pytest、PyInstaller。

验证：

- 安装依赖成功。
- `python -m pytest --collect-only` 成功。
- `python -m excel_accountant --help` 成功。

## 任务 2：领域模型和无损十进制编码

创建：

- `src/excel_accountant/models.py`
- `src/excel_accountant/decimal_codec.py`
- `tests/test_decimal_codec.py`

实现：

- 严格金额字符串解析。
- 尾随零与不同小数位的数学等价。
- 公共十进制尺度整数化。
- 最大公约数缩减。
- OR-Tools int64 安全范围检查。
- 金额、目标、方案、状态等不可变领域对象。

验证：

- 覆盖完整小数、负数、千位分隔符、非法文本和安全范围。
- 不使用浮点数作为金额真值。

## 任务 3：XLSX 精确读取和工作簿安全预检

创建：

- `src/excel_accountant/xlsx_reader.py`
- `src/excel_accountant/workbook_safety.py`
- `tests/test_xlsx_reader.py`
- `tests/test_workbook_safety.py`

实现：

- 解析 workbook relationships、sheet XML 和 shared strings。
- 从 XML `<v>` 读取原始数值文本。
- openpyxl 仅用于类型、日期、样式和结构辅助判断。
- 列字母、列序号、单列 A1 范围解析。
- 扫描截止到最后一个非空单元格。
- 公式、日期、布尔、错误、零值、隐藏行和异常文本分类。
- 检测宏、绘图、图表、嵌入对象、外部链接、保护和其他不安全结构。

验证：

- 测试工作簿由测试代码动态生成。
- 真实示例只用于本地端到端验证，不提交仓库。

## 任务 4：精确多目标互斥求解器

创建：

- `src/excel_accountant/solver_exact.py`
- `tests/test_solver_exact.py`

实现：

- `x[i,j]` 二元分配变量。
- 每个源单元格最多分配给一个目标。
- 每个目标非空且精确等于目标金额。
- 取消、时限和明确状态。
- no-good 约束枚举最多 100 套方案。
- 相同目标标签交换的等价方案规范化去重。

验证：

- 小规模随机数据使用暴力枚举真值对照。
- 覆盖重复金额、重复目标、负数、零目标、无解、多解、超时和取消。

## 任务 5：近似整体互斥求解器

创建：

- `src/excel_accountant/solver_approximate.py`
- `tests/test_solver_approximate.py`

实现：

- 仅在没有精确方案时调用。
- 保持源单元格互斥。
- 依次优化：精确目标数、总绝对差额、最大单项目标差额、使用单元格数。
- 返回最多 5 套，不提供导出能力。

验证：

- 使用可穷举小数据验证排序和差额。
- 区分“已证明无解”和“精确搜索超时”。

## 任务 6：安全输出与独立验证

创建：

- `src/excel_accountant/xlsx_writer.py`
- `src/excel_accountant/output_verifier.py`
- `tests/test_xlsx_output.py`

实现：

- 原文件哈希保护。
- 只对安全预检通过的普通 `.xlsx` 输出。
- 临时文件写入、验证和原子重命名。
- 仅高亮金额单元格。
- 新增“凑数结果”审计工作表。
- 每套方案独立文件。
- 验证失败删除临时文件。

验证：

- 重新读取每个输出，验证逐目标精确和、地址互斥、颜色和结果表一致。
- 验证原文件哈希不变。

## 任务 7：应用服务和后台执行

创建：

- `src/excel_accountant/service.py`
- `src/excel_accountant/worker.py`
- `tests/test_service.py`

实现：

- 串联读取、预览确认、编码、精确求解、近似求解、输出和验证。
- 状态枚举与中文用户消息一一映射。
- 进度、取消、超时和继续搜索接口。
- 精确方案已找到但输出被安全预检阻止时，不启动近似求解。

验证：

- 对关键状态进行服务层集成测试。

## 任务 8：PySide6 桌面界面

创建：

- `src/excel_accountant/__main__.py`
- `src/excel_accountant/ui/main_window.py`
- `src/excel_accountant/ui/models.py`

实现：

- 文件拖放和选择。
- 工作表、金额范围和目标金额表。
- 完整数值与异常数据预览。
- 方案数量 1 至 100、默认 20。
- 搜索时限默认 60 秒。
- 搜索、取消、继续和打开输出目录。
- 精确结果、近似结果、差额和状态展示。
- 近似结果可在风险确认后勾选导出，并记录实际合计与差额。

验证：

- GUI 冒烟测试。
- 后台搜索时界面保持响应。

## 任务 9：端到端验证、文档和打包

更新：

- `README.md`
- `scripts/build.ps1`

实现：

- 使用当前本地示例工作簿进行不提交数据的端到端测试。
- 先构建 one-folder，再构建 one-file Windows 包。
- 记录安装、运行、支持范围、隐私、错误状态和限制。

验证：

- 全量 pytest 通过。
- `git diff --check` 通过。
- one-folder 程序启动成功。
- one-file 程序启动成功。
- 输出文件可由本机 WPS 打开并通过独立验证。

## 任务 10：完成分支

- 复核全部测试和构建日志。
- 确认真实 Excel 文件未被 Git 跟踪。
- 提交功能分支并推送。
- 使用 `finishing-a-development-branch` 技能完成分支交付。
