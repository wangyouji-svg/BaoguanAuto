# 报关资料后端 - 服务器部署脚本
# 用法：在本地 PowerShell 执行：.\deploy.ps1
# 服务器：root@101.96.212.128

$SERVER = 'root@101.96.212.128'
$REMOTE_DIR = '/root/baoguan-backend'

Write-Host '=== 1. 上传文件 ===' -ForegroundColor Cyan
scp "$PSScriptRoot\backend_server.py" "${SERVER}:${REMOTE_DIR}/backend_server.py"
scp "$PSScriptRoot\..\报关资料模板.xlsx" "${SERVER}:${REMOTE_DIR}/报关资料模板.xlsx"

Write-Host '=== 2. 安装依赖并启动服务 ===' -ForegroundColor Cyan
ssh $SERVER @"
set -e
mkdir -p $REMOTE_DIR/generated

cd $REMOTE_DIR

# 安装依赖（如已安装会跳过）
pip install flask openpyxl -q

# 停掉旧进程（如果有）
pkill -f backend_server.py 2>/dev/null || true
sleep 1

# 后台启动
nohup python backend_server.py > $REMOTE_DIR/server.log 2>&1 &
sleep 2

# 验证是否启动成功
if curl -s -o /dev/null -w '%{http_code}' http://localhost:5000/generate -X POST -H 'Content-Type: application/json' -d '{"rows":[]}' | grep -q 200; then
    echo '✅ 服务启动成功'
else
    echo '❌ 服务启动失败，查看日志：'
    tail -20 $REMOTE_DIR/server.log
fi
"@

Write-Host '=== 完成 ===' -ForegroundColor Green
Write-Host "后端地址：http://101.96.212.128:5000/generate" -ForegroundColor Yellow
Write-Host "请确认服务器防火墙已开放 5000 端口" -ForegroundColor Yellow
