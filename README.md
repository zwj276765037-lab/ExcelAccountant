# ExcelAccountant

ExcelAccountant 是一个完全离线的 Windows 桌面工具：从 XLSX 的一列金额中，为一个或多个目标金额寻找精确组合。在同一套方案内，每个源单元格最多使用一次。

这是独立本地软件，不需要 WPS 插件环境。它生成的 `.xlsx` 可用 WPS Office 或 Microsoft Excel 打开。

## 金额准确性

- 直接从 XLSX 内部 XML 读取已保存的原始数值文本。
- 金额全程使用十进制定点整数化计算，不用浮点数作为求和真值。
- 不截断、不四舍五入、不主动统一小数位。
- 保证不在工作簿已保存的精度之上制造新的舍入；Excel/WPS 保存前已丢失的精度无法恢复。

## 使用方法

1. 运行 `ExcelAccountant.exe`。
2. 选择或拖入一个 `.xlsx` 文件，并选择工作表。
3. 输入金额列：列字母 `E`、列序号 `5`，或单列范围 `E2:E500`。
4. 点击“预览读取结果”，核对有效数据、完整小数和被排除单元格。
5. 在目标区每行输入一个金额。多个目标会在同一个输出文件中使用不同颜色。
6. 设置最多方案数和搜索时限，点击“开始精确搜索”。
7. 从“搜索结果”查看精确组合或近似备选。搜索完成时不会自动生成文件。
8. 在精确方案左侧勾选需要的方案，确认输出目录后点击“输出勾选方案”。

只有当所有目标在同一套方案中都精确满足，且用户手动勾选并点击输出时，程序才会生成文件。无精确方案时，程序仅显示最多 5 套互不重复的近似备选，不会导出近似结果。

## 输出文件

- 每套完整方案是一个独立文件：`原文件名_方案001.xlsx`、`原文件名_方案002.xlsx` 等。
- 同一文件中，不同目标使用不同填充色；只改变被选中的金额单元格。
- 新增“凑数结果”工作表，列出目标、组成金额、单元格地址、精确合计和复核状态。
- 原始文件永不覆盖。结果先写入临时文件，重新打开并独立验证通过后才落盘。
- 同一套方案内源单元格不重复；不同方案文件是相互独立的备选，允许复用相同源项目。

## 数据识别规则

作为候选金额：数值单元格、严格数值文本、正数、负数和隐藏行中的金额。

不参与搜索并在预览中标明：表头、空白、零值、公式、日期、布尔值、错误值、币种符号或混合文本。非表头异常项存在时，搜索前会再次请求确认。

## 支持范围与限制

第一版支持 Windows 10/11、普通 `.xlsx`、单工作表的单列范围、正负金额和数百条候选数据。

不支持 `.xls`、`.xlsm`、加密工作簿、公式金额、跨表/跨列共同搜索。宏、图表、形状、嵌入对象、外部链接、连接或保护等结构会触发安全拒绝：可以显示求解结果，但不会写回副本。

“任意数量数值求和”属于组合搜索问题，数据越多、目标越多、可选组合越密集，耗时可能越长。达到时限但没找到方案时，状态为“在当前时限内未找到”，不会误报为已证明无解。

## 隐私

程序完全在本机运行，无登录、无上传、无遥测、无云端 API。真实工作簿、`source/`、`output/` 及所有 Excel 输入/输出都被 Git 忽略，不会提交到公开仓库。

## 开发与测试

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m pip install -e .
.\.venv\Scripts\python.exe -m pytest -q
.\.venv\Scripts\python.exe -m excel_accountant
```

实现规格与测试计划见：

- `docs/superpowers/specs/2026-09-02-excel-accountant-design.md`
- `docs/superpowers/plans/2026-09-02-excel-accountant-implementation-plan.md`

## Windows 打包

```powershell
.\scripts\build.ps1
```

构建脚本会先运行全部测试，再生成并冒烟检查两种包：

- one-folder：`dist\ExcelAccountant-folder\ExcelAccountant-folder.exe`
- one-file：`dist\ExcelAccountant.exe`

仅在已单独运行测试时，可用 `.\scripts\build.ps1 -SkipTests` 跳过构建前测试。
