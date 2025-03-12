############################################################################################
# Q：これは何？
# A：Microsoft Edge上でIEモードで開くように設定したサイトが
#    30日経過後に有効期限切れで解除されるのがウザくて作ったやつ
#
# Usage：最初だけ → 同ディレクトリにある「GetIETabListName.bat」を実行
#        ①このシェルと同じディレクトリに、httpsなしのURLを記載したファイルを配置
#          ファイル名は「OpenIEUrlList.txt」、文字コードは「UTF-8」とすること
#        ②この処理を起動するといいかんじに設定される（はず）管理者権限で起動すること
# Tips ：タスクスケジューラーに管理者権限込みで登録して定期的に起動させるようにすると便利
############################################################################################
$LocalFlg = 0
$OldInternetExplorerIntegrationSiteListUrl = Get-Content InternetExplorerIntegrationSiteList.ini | Select-String  "http"
$dt = Get-Date -Format "yyyymmdd"
$dttm = Get-Date -Format "yyyymmdd_hhmmss"
$OldFilePath = "OldOpenIEUrlList.xml"
$OutputFilePath = "OpenIEUrlList.xml"

#文字化け防止
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'

# パス周りの処理
# フルパス起動とかの場合、自分自身のディレクトリに移動
$scriptPath = $MyInvocation.MyCommand.Path
$scriptParentPath = Split-Path -Parent $scriptPath
Set-Location -Path $scriptParentPath

# レジストリにIEモードの設定ファイルがセットされている場合
# そのファイルをダウンロードしてローカルに配置する
If($LocalFlg -eq 0){
    If((Test-Path Variable:OldInternetExplorerIntegrationSiteListUrl) -eq "TRUE"){
        try{
            Invoke-DownloadFile -Url $OldInternetExplorerIntegrationSiteListUrl -DownloadFolder $scriptParentPath -FileName $OldFilePath
        } catch {
            Write-Output $dttm + ' Download Failed' > errorlog.txt
        }
    }
}  

# ヘッダーを追加
Write-Output ('<site-list_version="' + $dt + '">') > $OutputFilePath

# 既にあるリストがあれば新規ファイルにセットする
if((Test-Path $OldFilePath) -eq "True"){
    Get-Content $OldFilePath | Select-String -NotMacth "site-list" >> $OutputFilePath
}

# データファイルのImport
$openWithIeSites = Get-Content OpenIEUrlList.txt -Encoding UTF8

ForEach($site in $openWithIeSites) {
    Write-Output ('  <site url="' + $site + '">') >> $OutputFilePath
    Write-Output '    <open-in>IE11</open-in>' >> $OutputFilePath
    Write-Output '  </site>' >> $OutputFilePath
}

#フッター追加
Write-Output '</site-list>' >>  $OutputFilePath


#とりあえずレジストリを更新する処理を入れる
try {
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "InternetExplorerIntegrationSiteList" -Value $scriptParentPath + '\' + $OutputFilePath
}
catch {
    Write-Output $dttm + ' RegistorySet Failed' > errorlog.txt
}