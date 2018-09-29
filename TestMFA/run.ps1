<## POST method: $req
$requestBody = Get-Content $req -Raw | ConvertFrom-Json
$name = $requestBody.name

# GET method: each querystring parameter is its own variable
if ($req_query_name) 
{
    $name = $req_query_name 
}

Out-File -Encoding Ascii -FilePath $res -inputObject "Hello $name"#>

Write-Output �Getting PowerShell Module�
$u = $ENV:user
$p = $ENV:password
Import-Module "D:\home\site\wwwroot\TestMFA\PSModules\MSOnline\1.1.183.17\MSOnline.psd1"
Import-Module "D:\home\site\wwwroot\TestMFA\PSModules\AzureAD\2.0.1.16\AzureAD.psd1"
$result = Get-Module -ListAvailable | Select-Object Name, Version, ModuleBase | Sort-Object -Property Name | Format-Table -wrap | Out-String

Write-output "User:$($u) Password: $($p) $($result)"