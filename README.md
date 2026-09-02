# ExcelAccountant

ExcelAccountant 是一个完全离线的 Windows 桌面工具，用于从 XLSX 单列金额中寻找多个目标金额的互斥精确组合。

当前开发分支依据以下规格实施：

- `docs/superpowers/specs/2026-09-02-excel-accountant-design.md`
- `docs/superpowers/plans/2026-09-02-excel-accountant-implementation-plan.md`

真实财务工作簿、输入文件和输出结果均被 Git 忽略，不会提交到公开仓库。

## 开发环境

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m pip install -e .
.\.venv\Scripts\python.exe -m pytest
```

## 运行

```powershell
.\.venv\Scripts\python.exe -m excel_accountant
```

## 打包

```powershell
.\scripts\build.ps1
```

支持范围、操作说明和发行包信息将在功能完成后补充。
