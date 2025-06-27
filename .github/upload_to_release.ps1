# 配置日志文件路径
$desktopPath = [Environment]::GetFolderPath("Desktop")
$logFilePath = "$desktopPath\release_log.txt"
Start-Transcript -Path $logFilePath

try {
    # 从环境变量中获取 GitHub Token
    $token = [System.Environment]::GetEnvironmentVariable("GITHUB_TOKEN")

    # 检查 Token 是否已设置
    if (-not $token) {
        Write-Error "GitHub Token 未设置！请将 Token 存储为环境变量 GITHUB_TOKEN。"
        exit
    }

    Write-Output "当前 GitHub Token 为: $token"

    # 配置变量
    $repo = "defghdffg/dwy"  # 仓库路径
    $deviceName = $env:COMPUTERNAME  # 获取设备名称
    $tagName = "v1.0.0"  # Release 标签
    $releaseName = "Initial Release - $deviceName"  # Release 名称
    $releaseDescription = "包含来自设备 [$deviceName] 的快捷方式路径和批量上传的文件压缩包"  # Release 描述
    $shortcutFile = "$desktopPath\shortcut_paths_$deviceName.txt"  # 保存快捷方式路径的文件
    $shortcutZip = "$desktopPath\shortcut_paths_$deviceName.zip"  # 快捷方式路径压缩包
    $batchUploadFolder = "C:\Users\Administrator\Documents\MyFiles"  # 文件夹路径
    $batchUploadZip = "$desktopPath\batch_uploads_$deviceName.zip"  # 文件夹压缩包

    # 检查路径是否有效
    if (-not (Test-Path $batchUploadFolder)) {
        Write-Error "文件夹路径不存在: $batchUploadFolder"
        exit
    }
    if (-not (Test-Path $desktopPath)) {
        Write-Error "桌面路径不存在: $desktopPath"
        exit
    }

    # 提取桌面快捷方式路径
    $shell = New-Object -ComObject WScript.Shell
    Write-Output "提取桌面快捷方式路径中..."
    $shortcutPaths = @()
    Get-ChildItem -Path $desktopPath -Filter *.lnk | ForEach-Object {
        try {
            $shortcut = $shell.CreateShortcut($_.FullName)
            $shortcutPaths += "$($_.Name): $($shortcut.TargetPath)"
        } catch {
            Write-Warning "无法读取快捷方式: $($_.FullName)"
        }
    }

    if ($shortcutPaths.Count -eq 0) {
        Write-Error "未找到有效的快捷方式！"
        exit
    }

    # 保存快捷方式路径到文本文件
    $shortcutPaths | Out-File -FilePath $shortcutFile -Encoding UTF8

    # 压缩快捷方式路径文件
    if (Test-Path $shortcutZip) { Remove-Item -Path $shortcutZip }
    Compress-Archive -Path $shortcutFile -DestinationPath $shortcutZip

    # 压缩批量上传文件
    if (-not (Get-ChildItem -Path "$batchUploadFolder\*" -ErrorAction SilentlyContinue)) {
        Write-Error "批量上传文件夹为空: $batchUploadFolder"
        exit
    }
    if (Test-Path $batchUploadZip) { Remove-Item -Path $batchUploadZip }
    Compress-Archive -Path "$batchUploadFolder\*" -DestinationPath $batchUploadZip

    # 创建 Release
    Write-Output "创建 Release..."
    $releaseUrl = "https://api.github.com/repos/$repo/releases"
    $headers = @{
        Authorization = "Bearer $token"
        Accept = "application/vnd.github.v3+json"
    }
    $body = @{
        tag_name = $tagName
        name = $releaseName
        body = $releaseDescription
        draft = $false
        prerelease = $false
    } | ConvertTo-Json -Depth 10

    $response = Invoke-RestMethod -Uri $releaseUrl -Method Post -Headers $headers -Body $body
    $releaseId = $response.id
    Write-Output "Release 创建成功，ID: $releaseId"

    # 上传附件函数
    function Upload-Asset {
        param (
            [string]$filePath,
            [string]$releaseId
        )
        $fileName = [System.IO.Path]::GetFileName($filePath)
        $uploadUrl = "https://uploads.github.com/repos/$repo/releases/$releaseId/assets?name=$fileName"
        $headers = @{
            Authorization = "Bearer $token"
            "Content-Type" = "application/zip"
        }
        $response = Invoke-RestMethod -Uri $uploadUrl -Method Post -Headers $headers -InFile $filePath -ContentType "application/zip"
        Write-Output "文件已上传: $($response.browser_download_url)"
    }

    # 上传压缩包
    Upload-Asset -filePath $shortcutZip -releaseId $releaseId
    Upload-Asset -filePath $batchUploadZip -releaseId $releaseId

    Write-Output "所有任务完成！"
} catch {
    Write-Error "发生错误：$($_.Exception.Message)"
} finally {
    Stop-Transcript
    Write-Output "执行日志已保存到: $logFilePath"
}
