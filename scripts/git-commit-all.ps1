Param(
	[string]$Message,
	[switch]$Push,
	[string]$Remote = 'origin',
	[switch]$NoVerify
)

function ExitWith($code, $msg) {
    if ($msg) { Write-Host $msg }
    exit $code
}

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    ExitWith 1 "未找到 git，请先安装并配置 PATH。"
}

if (-not $Message) {
    $Message = Read-Host "请输入提交信息"
}
if (-not $Message) {
    ExitWith 1 "已取消：提交信息为空。"
}

Write-Host "添加所有更改..."
& git add -A
if ($LASTEXITCODE -ne 0) {
    ExitWith 1 "git add 失败。"
}

$commitArgs = @('commit', '-m', $Message)
if ($NoVerify) { $commitArgs += '--no-verify' }

Write-Host "正在提交..."
& git @commitArgs
if ($LASTEXITCODE -ne 0) {
    Write-Host "git commit 返回非零，可能无更改可提交或发生错误。"
    exit $LASTEXITCODE
}

Write-Host "提交成功。"

if ($Push) {
    $branch = (& git rev-parse --abbrev-ref HEAD).Trim()
    if (-not $branch) {
        ExitWith 1 "无法获取当前分支名。"
    }
    Write-Host "正在推送到 $Remote/$branch ..."
    & git push $Remote $branch
    if ($LASTEXITCODE -ne 0) {
        ExitWith 1 "git push 失败。"
    }
    Write-Host "推送成功。"
} else {
    Write-Host "已完成提交（未推送）。使用 -Push 参数可自动推送。"
}