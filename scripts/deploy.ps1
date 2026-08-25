# 报关资料后端 - 服务器部署脚本
# 用法：在本地 PowerShell 执行：.\deploy.ps1

$ErrorActionPreference = 'Stop'
$Server = 'root@101.96.212.128'
$RemoteDir = '/youji/apps/baoguan-backend'
$StageDir = "/tmp/baoguan-deploy-$([DateTimeOffset]::Now.ToUnixTimeSeconds())"
$Python = '/youji/apps/logiflow-tracker/.venv/bin/python'

$Files = @(
    'backend_server.py',
    'mysql_database.py',
    'migrate_to_mysql.py',
    'baoguan.service',
    'dingtalk_demo.js',
    'test_regression.py',
    'test_mysql_storage.py',
    'test_dingtalk_batch.js'
)

Write-Host '1. 上传版本文件到临时目录' -ForegroundColor Cyan
ssh $Server "mkdir -p '$StageDir/mysql_migrations'"
foreach ($File in $Files) {
    scp (Join-Path $PSScriptRoot $File) "${Server}:${StageDir}/$File"
}
scp (Join-Path $PSScriptRoot 'mysql_migrations\*.sql') "${Server}:${StageDir}/mysql_migrations/"
scp (Join-Path $PSScriptRoot '..\requirements.txt') "${Server}:${StageDir}/requirements.txt"
scp (Join-Path $PSScriptRoot '..\README.md') "${Server}:${StageDir}/README.md"
scp (Join-Path $PSScriptRoot '..\CHANGELOG.md') "${Server}:${StageDir}/CHANGELOG.md"

Write-Host '2. 安装依赖、切换文件并重启 systemd 服务' -ForegroundColor Cyan
ssh $Server @"
set -eu
test -f '$RemoteDir/.env.mysql'
backup='$RemoteDir/backups/deploy-'`$(date +%Y%m%d-%H%M%S)
mkdir -p "`$backup" '$RemoteDir/mysql_migrations' '$RemoteDir/generated'
cp -a '$RemoteDir/backend_server.py' '$RemoteDir/mysql_database.py' '$RemoteDir/baoguan.service' "`$backup/" 2>/dev/null || true
'$Python' -m pip install -r '$StageDir/requirements.txt' -q
cp -a '$StageDir/.' '$RemoteDir/'
install -m 0644 '$RemoteDir/baoguan.service' /etc/systemd/system/baoguan.service
systemctl daemon-reload
systemctl restart baoguan
for attempt in 1 2 3 4 5; do
    if curl --fail --silent http://127.0.0.1:5000/health | grep -q '"storageBackend":"mysql"'; then
        systemctl is-active --quiet baoguan
        rm -rf '$StageDir'
        exit 0
    fi
    sleep 2
done
journalctl -u baoguan -n 50 --no-pager
exit 1
"@

Write-Host '3. MySQL 健康检查通过，部署完成' -ForegroundColor Green
Write-Host '地址：https://pkcellsolution.com/baoguan/generate' -ForegroundColor Yellow
Write-Host '说明：.env.mysql 只保存在服务器，脚本不会上传或覆盖数据库密钥。' -ForegroundColor Yellow
