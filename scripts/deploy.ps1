# Customs document backend deployment.
# Run locally with: .\deploy.ps1

$ErrorActionPreference = 'Stop'
$Server = 'root@101.96.212.128'
$RemoteDir = '/youji/apps/baoguan-backend'
$StageDir = "/tmp/baoguan-deploy-$([DateTimeOffset]::Now.ToUnixTimeSeconds())"
$Python = '/youji/apps/logiflow-tracker/.venv/bin/python'
$Commit = (git -C (Join-Path $PSScriptRoot '..') rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0) { throw 'Failed to resolve the local Git commit.' }

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

Write-Host '1. Uploading release files to staging' -ForegroundColor Cyan
ssh $Server "mkdir -p '$StageDir/mysql_migrations'"
if ($LASTEXITCODE -ne 0) { throw 'Failed to create remote staging directory.' }
foreach ($File in $Files) {
    scp (Join-Path $PSScriptRoot $File) "${Server}:${StageDir}/$File"
    if ($LASTEXITCODE -ne 0) { throw "Failed to upload $File." }
}
scp (Join-Path $PSScriptRoot 'mysql_migrations\*.sql') "${Server}:${StageDir}/mysql_migrations/"
if ($LASTEXITCODE -ne 0) { throw 'Failed to upload MySQL migrations.' }
scp (Join-Path $PSScriptRoot '..\requirements.txt') "${Server}:${StageDir}/requirements.txt"
if ($LASTEXITCODE -ne 0) { throw 'Failed to upload requirements.txt.' }
scp (Join-Path $PSScriptRoot '..\README.md') "${Server}:${StageDir}/README.md"
if ($LASTEXITCODE -ne 0) { throw 'Failed to upload README.md.' }
scp (Join-Path $PSScriptRoot '..\CHANGELOG.md') "${Server}:${StageDir}/CHANGELOG.md"
if ($LASTEXITCODE -ne 0) { throw 'Failed to upload CHANGELOG.md.' }

Write-Host '2. Installing dependencies and restarting systemd service' -ForegroundColor Cyan
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
for attempt in 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15; do
    if curl --fail --silent http://127.0.0.1:5000/health | grep -q mysql; then
        systemctl is-active --quiet baoguan
        printf '%s\n' '$Commit' > '$RemoteDir/DEPLOYED_COMMIT'
        rm -rf '$StageDir'
        exit 0
    fi
    sleep 2
done
journalctl -u baoguan -n 50 --no-pager
exit 1
"@
if ($LASTEXITCODE -ne 0) { throw 'Remote deployment or health check failed.' }

Write-Host '3. MySQL health check passed; deployment completed' -ForegroundColor Green
Write-Host 'URL: https://pkcellsolution.com/baoguan/generate' -ForegroundColor Yellow
Write-Host 'The server-only .env.mysql file is never uploaded or overwritten.' -ForegroundColor Yellow
