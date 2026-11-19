@echo off
chcp 936 >nul 2>nul
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================================
echo   结算报表自动化引擎 - 一键运行
echo ============================================================
echo.

rem ---------- 1. 查找 Python ----------
set "PYTHON="
where python >nul 2>nul
if not errorlevel 1 (
    python -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
    if not errorlevel 1 set "PYTHON=python"
)
if not defined PYTHON (
    where py >nul 2>nul
    if not errorlevel 1 (
        py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>nul
        if not errorlevel 1 set "PYTHON=py -3"
    )
)
if not defined PYTHON (
    echo [错误] 未找到 Python 3.10 或更高版本。
    echo        请先安装 Python，并在安装时勾选 "Add Python to PATH"。
    echo        下载地址: https://www.python.org/downloads/
    goto :end
)
echo [信息] 使用 Python: %PYTHON%
%PYTHON% -c "import sys; print('Python', sys.version.split()[0])"
if errorlevel 1 goto :error

rem ---------- 2. 检查依赖 ----------
echo.
echo [1/4] 检查运行依赖...
%PYTHON% -c "import openpyxl, docx, docxtpl" >nul 2>nul
if errorlevel 1 (
    echo [提示] 缺少依赖，正在安装 requirements.txt ...
    if not exist "requirements.txt" (
        echo [错误] 未找到 requirements.txt，无法自动安装依赖。
        goto :error
    )
    %PYTHON% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [错误] 依赖安装失败。请检查网络后重试，或手动执行:
        echo        %PYTHON% -m pip install -r requirements.txt
        goto :error
    )
)
echo [完成] 依赖检查通过。

rem ---------- 3. 初始化配置 ----------
echo.
echo [2/4] 检查规则配置...
if not exist "规则配置.xlsx" (
    echo [提示] 未找到 规则配置.xlsx，正在生成初始配置...
    %PYTHON% init_config.py
    if errorlevel 1 goto :error
) else (
    echo [信息] 已找到 规则配置.xlsx，跳过初始化。
)

rem ---------- 4. 生成 Word 内部模板 ----------
echo.
echo [3/4] 检查原版表格并生成模板...
if exist "原版表格.docx" (
    %PYTHON% make_templates.py
    if errorlevel 1 goto :error
) else (
    echo [警告] 未找到 原版表格.docx，跳过模板生成。
    if not exist "_内部模板.docx" (
        echo [错误] 也没有找到 _内部模板.docx，无法继续。
        echo        请把原始 Word 表格命名为 原版表格.docx 放进本文件夹。
        goto :error
    )
    if not exist "_模板结构.json" (
        echo [错误] 找到 _内部模板.docx，但缺少 _模板结构.json。
        echo        请把 原版表格.docx 放回本文件夹，重新运行一次以生成模板结构。
        goto :error
    )
    echo [信息] 将使用已有的 _内部模板.docx。
)

rem ---------- 5. 生成报告 ----------
echo.
echo [4/4] 读取源数据并生成报告...
if not exist "源数据.xlsx" (
    echo [错误] 未找到 源数据.xlsx。
    echo        请把填写好的数据保存为 源数据.xlsx 放进本文件夹。
    echo        首次使用可参考 源数据模板.xlsx。
    goto :error
)
%PYTHON% make_report.py
if errorlevel 1 goto :error

echo.
echo ============================================================
echo   全部完成。请查看生成的 结算报表_日期.docx
echo ============================================================
goto :end

:error
echo.
echo ============================================================
echo   运行失败，请根据上面的错误提示处理。
echo ============================================================

:end
echo.
pause
endlocal
