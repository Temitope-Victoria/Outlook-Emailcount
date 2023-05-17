Get-MailboxFolderStatistics temitope.opeagbe@funmilayo.tk | Select Name,FolderSize,ItemsinFolder

$outlook = new-object -com Outlook.Application
$session = $outlook.Session
$session.Logon()
$inbox = $outlook.session.GetDefaultFolder(6)
[array]$unreadCount = @(%{$inbox.Items | where {$_.UnRead}}).Count
Write-Host $unreadCount
[array]$ItemCount = @(%{$inbox.Items }).Count
Write-Host $ItemCount