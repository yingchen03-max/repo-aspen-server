@echo off
chcp 65001 >nul
echo ============================================================
echo Aspen Agent 服务启动脚本
echo ============================================================
echo.

REM 检查 Python 是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.8+
    pause
    exit /b 1
)

REM 检查是否存在 .env 文件
if not exist .env (
    echo [警告] 未找到 .env 文件
    echo [提示] 正在从 .env.example 创建 .env 文件...
    if exist .env.example (
        copy .env.example .env
        echo [成功] 已创建 .env 文件，请编辑此文件配置您的环境
        echo.
        pause
    ) else (
        echo [错误] 未找到 .env.example 文件
        pause
        exit /b 1
    )
)

REM 检查依赖是否安装
echo [检查] 正在检查 Python 依赖...
python -c "import flask" >nul 2>&1
if errorlevel 1 (
    echo [警告] 未安装依赖包
    echo [提示] 正在安装依赖...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [错误] 依赖安装失败
        pause
        exit /b 1
    )
)

echo [成功] 依赖检查完成
echo.

REM 验证配置
echo [检查] 正在验证配置...
python config.py
echo.

REM 准备固定 SSL 证书
echo [SSL] 正在检查并准备 HTTPS 证书...
python bootstrap_ssl.py
if errorlevel 1 (
    echo [错误] SSL 证书准备失败，服务不会以 HTTP 回退启动
    pause
    exit /b 1
)
echo.

REM 启动服务
echo [启动] 正在启动 Aspen Agent 服务...
echo.
python aspenagent.py

pause
