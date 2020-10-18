UpdateWorkDays.psm1 は Exchange サーバー メールボックスの予定表の "WorkDays" の設定を修正するための PowerShell スクリプトです。  
このファイル内で `Update-WorkDays` と `Get-Token` 関数が定義されています。

[ダウンロード](https://github.com/jpmessaging/UpdateWorkDays/releases/download/v2020-10-18/UpdateWorkDays.zip)

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
- PowerShell v3 以上
- .NET Framework 4.6.1 以上

下記モジュールは本スクリプトのあるパスの "modules" フォルダ内に配置されている必要があります。これらはリリース パッケージに含まれています。

- [Microsoft Exchange Web Services Managed API 2.2](https://www.microsoft.com/en-us/download/details.aspx?id=42951)
- [MSAL.NET](https://www.nuget.org/packages/Microsoft.Identity.Client)
- [MSAL.NET Extensions](https://www.nuget.org/packages/Microsoft.Identity.Client.Extensions.Msal/)

## 利用方法
1. Exchange 管理シェルにて、"ApplicationImpersonation" の役割をスクリプトを実行するユーザーに付与します

   これは本スクリプトでは 偽装 (Impersonation) を使って対象のメールボックスにアクセスするためです。

   ```PowerShell
   New-ManagementRoleAssignment -Role "ApplicationImpersonation" -User contoso\administrator
   ```

2. PowerShell を開始して、Import-Module で本スクリプトを取り込みます

   例:
   ```PowerShell
   Import-Module 'C:\tmp\UpdateWorkDays.psm1'
   ```

3. 先進認証を利用する場合には利用するアクセス トークンを取得します。

   トークン取得には本スクリプトに含まれる `Get-Token` を利用することができます。

   例:
   ```PowerShell
   $token = Get-Token -ClientId '63ce5cc6-c944-4baa-83d1-5cac8cdf487e' -Scopes 'https://outlook.office365.com/EWS.AccessAsUser.All'
   ```

4. Update-WorkDays を実行します

   #### 必須パラメータ:

   | name                     | meaning                                                 |
   | ------------------------ | ------------------------------------------------------- |
   | Server                   | EWS サーバー名。EXO の場合には `outlook.office365.com`  |
   | TargetMailboxSmtpAddress | WorkDays を更新する対象のメールボックスの SMTP アドレス |


   #### 条件付き必須パラメータ:
   ※ 下記のパラメータは排他的です。レガシー認証を利用する場合には `Credential` を指定し、先進認証を利用する場合は `Token` を指定します。

   | name       | meaning                                                              |
   | ---------- | -------------------------------------------------------------------- |
   | Credential | レガシー認証利用時に対象のメールボックスへアクセスするための資格情報 |
   | Token      | 先進認証を利用時に利用するアクセス トークン                          |

   #### オプショナル パラメータ:

   | name              | meaning                                      |
   | ----------------- | -------------------------------------------- |
   | EwsManagedApiPath | Microsoft.Exchange.WebServices.dll のパス    |
   | EnableTrace       | トレースの出力を有効化する Switch パラメータ |
   | TraceFile         | トレースを出力する先のファイルのパス         |


   例 1: オンプレミスのメールボックスに対して実行
   ```PowerShell
   Update-WorkDays -Server 'myExchange.contoso.local' -Credential (Get-Credential) -TargetMailboxSmtpAddress 'user01@contoso.local'
   ```

   例 2: 先進認証を利用して EXO メールボックスに対して実行     
   ※ 下記の Client ID はあくまでサンプルです。当該環境で登録したアプリケーションの Client ID を利用ください。  
   ※ 以下では、アプリケーションがマルチ テナント アプリケーションという前提のため、`Get-Token` の `TenantId` パラメータはスキップしています。シングル テナント アプリケーションの場合には、TenantId に対象テナント名や GUID を指定ください。  

   ```PowerShell
   $token = Get-Token -ClientId '63ce5cc6-c944-4baa-83d1-5cac8cdf487e' -Scopes 'https://outlook.office365.com/EWS.AccessAsUser.All'
   Update-WorkDays -Server 'outlook.office365.com' -TargetMailboxSmtpAddress 'room01@contoso.com' -Token $token.AccessToken -EnableTrace -TraceFile 'C:\temp\trace.txt'
   ```



## 先進認証について
先進認証を利用する場合には、事前に Azure AD にアプリケーションを登録する必要があります。

- `Get-Token` は既定の Redirect URI として `https://login.microsoftonline.com/common/oauth2/nativeclient` を利用します。異なる Redirect URI を登録する場合には、`Get-Token` の `RedirectUri` パラメータに指定ください。
- API Permissions には `Exchange` の `EWS.AccessAsUser.All` を付与します。

### 参考
- [Authenticate an EWS application by using OAuth](https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)

## ライセンス
Copyright (c) 2020 Ryusuke Fujita

This software is released under the MIT License.  
http://opensource.org/licenses/mit-license.php

以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「ソフトウェア」）の複製を取得するすべての人に対し、ソフトウェアを無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。

上記の著作権表示および本許諾表示を、ソフトウェアのすべての複製または重要な部分に記載するものとします。

ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての保証も含みますが、それに限定されるものではありません。 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わないものとします。



