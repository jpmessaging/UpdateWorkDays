# UpdateWorkDays
UpdateWorkDays.ps1 は Exchange サーバー メールボックスの予定表の "WorkDays" の設定を修正するための PowerShell スクリプトです。このファイル内で `Update-WorkDays` 関数が定義されています。

# Background
Exchange サーバーのコマンドレットで以下のように WorkDays を "Weekdays" や "AllDays" と構成することができますが、Outlook はこれらの値を理解しません。

e.g.
```PowerShell
Set-MailboxCalendarConfiguration user01 -WorkDays AllDays
```

その結果、予定表がグレイアウト (WorkDays がないような状態) となります。
本スクリプトは WorkDays の値を Outlook が理解できるように以下のとおり変更します。


- "AllDays" の場合  --> "Sunday Monday Tuesday Wednesday Thursday Friday Saturday"
- "Weekdays" の場合 --> "Monday Tuesday Wednesday Thursday Friday"

# Requirement
- PowerShell v2 or later
- [Microsoft Exchange Web Services Managed API 2.2](https://www.microsoft.com/en-us/download/details.aspx?id=42951)

# How to use

1. Exchange 管理シェルにて、"ApplicationImpersonation" の役割をスクリプトを実行するユーザーに付与します  
   これは本スクリプトでは 偽装 (Impersonation) を使って対象のメールボックスにアクセスするためです。

   ```PowerShell
   New-ManagementRoleAssignment -Role "ApplicationImpersonation" -User contoso\administrator
   ```
   
2. EWS Managed API をダウンロードしてスクリプトを実行するマシンにインストールします  
   実際には、Microsoft.Exchange.WebServices.dll があれば良いだけなので他のマシンにインストールして当該 DLL だけコピーしていただいても結構です
   
   [Microsoft Exchange Web Services Managed API 2.2](https://www.microsoft.com/en-us/download/details.aspx?id=42951)

3. PowerShell を開始して、ドット ソースで本スクリプトを取り込みます

   e.g. 
   ```
   . c:\tmp\UpdateWorkDays.ps1
   ```
  
5. Update-WorkDays を実行します

   必須パラメータ:
   
   |name|meagning
   |----|-
   |EwsManagedApiPath|Microsoft.Exchange.WebServices.dll のパス
   |Server|EWS サーバー名。EXO の場合には、outlook.office365.com または outlook.office.com 
   |Credential|対象のメールボックスへアクセスするための資格情報 (ApplicationImpersonation の役割をアサインされている必要あり)
   |TargetMailboxSmtpAddress|WorkDays を更新する対象のメールボックスの SMTP アドレスです。
   
   オプショナル パラメータ:

   |name|meagning
   |----|-
   |EnableTrace|トレースの出力を有効化する Switch パラメータ
   |TraceFile|トレースを出力する先のファイルのパス
   
   
   e.g.
   ```PowerShell
   Update-WorkDays -EwsManagedApiPath "C:\Microsoft.Exchange.WebServices.dll" -Server myExchange.contoso.local -Credential (Get-Credential) -TargetMailboxSmtpAddress user01@contoso.local
   ```
   
   
   
  
