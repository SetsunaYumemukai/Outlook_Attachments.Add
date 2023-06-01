# Outlookアプリケーションを作成
$Outlook = New-Object -ComObject Outlook.Application

# メールアイテムを作成
$Mail = $Outlook.CreateItem(0)

# 宛先、件名、本文を設定
$Mail.To = "example@example.com"
$Mail.Subject = "添付ファイル付きメール"
$Mail.Body = "このメールには添付ファイルがあります。"

# 添付ファイルのパスを設定
$AttachmentPath = "C:\path\to\file.txt"

# 添付ファイルを追加
$Attachment = $Mail.Attachments.Add($AttachmentPath)

# メールを表示（送信しない）
$Mail.Display()

# リソースを解放
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Attachment) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Mail) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
Remove-Variable Attachment, Mail, Outlook