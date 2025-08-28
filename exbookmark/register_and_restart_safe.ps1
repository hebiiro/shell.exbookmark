# register_and_restart_safe.ps1
# 管理者権限で実行すること（Visual Studioも管理者モードで起動）
param (
    [string]$dllOutputPath,
    [string]$dllTargetPath
)

function Close-ProcessesUsingDll($dllPath) {
    Write-Host "=== DLLを使用中のプロセスを検索中: $dllPath ==="
    $procs = Get-Process | ForEach-Object {
        try {
            $_.Modules | Where-Object { $_.FileName -ieq $dllPath } | ForEach-Object { $_.ModuleName; $_.BaseAddress; $_.FileName; $_ }
            if ($_.Modules.FileName -contains $dllPath) { $_ }
        } catch {
            # アクセス拒否は無視（別ユーザー権限プロセスなど）
        }
    }

    foreach ($p in $procs | Sort-Object -Unique) {
        Write-Host "プロセス終了: $($p.ProcessName) (PID: $($p.Id))"
        try {
            Stop-Process -Id $p.Id -Force
        } catch {
            Write-Warning "終了できませんでした: $($p.ProcessName)"
        }
    }
}

Write-Host "=== DLLを使用している全プロセスを終了します ==="
Close-ProcessesUsingDll -dllPath $dllTargetPath

Start-Sleep -Seconds 1

Write-Host "=== DLLをコピーします ==="
Copy-Item $dllOutputPath $dllTargetPath -Force

Write-Host "=== DLLを登録します ==="
& regsvr32.exe /s $dllTargetPath

# Write-Host "=== Explorer を再起動します ==="
# Start-Process explorer.exe

Write-Host "=== 完了しました ==="
