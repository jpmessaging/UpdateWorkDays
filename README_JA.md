UpdateWorkDays.ps1 は Exchange サーバー メールボックスの予定表の "WorkDays" の設定を修正するための PowerShell スクリプトです。このファイル内で `Update-WorkDays` 関数が定義されています。

[ダウンロード](https://github.com/jpmessaging/UpdateWorkDays/releases/download/v1.0/UpdateWorkDays.ps1)

## 背景
Exchange サーバーのコマンドレットで以下のように WorkDays を "Weekdays" や "AllDays" と構成することができますが、Outlook はこれらの値を理解しません。

e.g.
```PowerShell
Set-MailboxCalendarConfiguration user01 -WorkDays AllDays
```

その結果、予定表がグレイアウト (WorkDays がないような状態) となります。
本スクリプトは WorkDays の値を Outlook が理解できるように以下のとおり変更します。


- "AllDays" の場合  --> "Sunday Monday Tuesday Wednesday Thursday Friday Saturday"
- "Weekdays" の場合 --> "Monday Tuesday Wednesday Thursday Friday"

## 前提条件
- PowerShell v2 or later
- [Microsoft Exchange Web Services Managed API 2.2](https://www.microsoft.com/en-us/download/details.aspx?id=42951)

## 利用方法
1. Exchange 管理シェルにて、"ApplicationImpersonation" の役割をスクリプトを実行するユーザーに付与します  
   これは本スクリプトでは 偽装 (Impersonation) を使って対象のメールボックスにアクセスするためです。

   ```PowerShell
   New-ManagementRoleAssignment -Role "ApplicationImpersonation" -User contoso\administrator
   ```
   
2. EWS Managed API をダウンロードしてスクリプトを実行するマシンにインストールします  
   実際には、Microsoft.Exchange.WebServices.dll があれば良いだけなので他のマシンにインストールして当該 DLL だけコピーしていただいても結構です
   
   [Microsoft Exchange Web Services Managed API 2.2](https://www.microsoft.com/en-us/download/details.aspx?id=42951)

3. PowerShell を開始して、Import-Module で本スクリプトを取り込みます

   e.g. 
   ```
   Import-Module C:\tmp\UpdateWorkDays.ps1
   ```
  
5. Update-WorkDays を実行します

   必須パラメータ:
   
   | name                     | meaning                                                                                                        |
   | ------------------------ | -------------------------------------------------------------------------------------------------------------- |
   | EwsManagedApiPath        | Microsoft.Exchange.WebServices.dll のパス                                                                      |
   | Server                   | EWS サーバー名。EXO の場合には、outlook.office365.com または outlook.office.com                                |
   | Credential               | 対象のメールボックスへアクセスするための資格情報 (ApplicationImpersonation の役割をアサインされている必要あり) |
   | TargetMailboxSmtpAddress | WorkDays を更新する対象のメールボックスの SMTP アドレスです。                                                  |
   
   オプショナル パラメータ:

   | name        | meaning                                      |
   | ----------- | -------------------------------------------- |
   | EnableTrace | トレースの出力を有効化する Switch パラメータ |
   | TraceFile   | トレースを出力する先のファイルのパス         |
   
   
   e.g.
   ```PowerShell
   Update-WorkDays -EwsManagedApiPath "C:\Microsoft.Exchange.WebServices.dll" -Server myExchange.contoso.local -Credential (Get-Credential) -TargetMailboxSmtpAddress user01@contoso.local
   ```

## ライセンス
Copyright (c) 2020 Ryusuke Fujita

This software is released under the MIT License.  
http://opensource.org/licenses/mit-license.php

以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「ソフトウェア」）の複製を取得するすべての人に対し、ソフトウェアを無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。

上記の著作権表示および本許諾表示を、ソフトウェアのすべての複製または重要な部分に記載するものとします。

ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての保証も含みますが、それに限定されるものではありません。 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わないものとします。   
   
   
  
