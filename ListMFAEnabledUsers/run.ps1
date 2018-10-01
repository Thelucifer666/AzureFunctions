Write-Output "Function started execution at $(Get-Date) "
Write-Output "Gathering credentials"
$Timer =  [system.diagnostics.stopwatch]::StartNew()
try{
    $username = $Env:user
    $pw = $Env:password
    $keypath = "D:\home\site\wwwroot\ListMFAEnableUser\PassEncryptKey.key"
    $secpassword = $pw | ConvertTo-SecureString -Key (Get-Content $keypath)
    $credential = New-Object System.Management.Automation.PSCredential ($username, $secpassword)
}
catch{
    $Out = "Creating the credential object failed.
    $($_.Invocationinfo.MyCommand) at position $($_.Invocationinfo.positionmessage) failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
    $Timer.Stop()
    Out-File -Encoding Ascii -FilePath $res -inputObject $Out
    Return
}
try{
    Write-Output "Importing MSOnline PS Module"
    Import-Module "D:\home\site\wwwroot\ListMFAEnabledUsers\bin\MSOnline\1.1.183.17\MSOnline.psd1"
    Write-Output "Imported MSOnline PS Module"
    Write-Output "Importing AzureAD PS Module"
    Import-Module "D:\home\site\wwwroot\ListMFAEnabledUsers\bin\AzureAD\2.0.1.16\AzureAD.psd1"
    Write-Output "Imported MSOnline PS Module"
}
catch{
    $Out = "Importing module failed.
    $($_.Invocationinfo.MyCommand) at position $($_.Invocationinfo.positionmessage) failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
    $Timer.Stop()
    Out-File -Encoding Ascii -FilePath $res -inputObject $Out
    Return
}
# Connect to MSOnline
try{
    Write-Output "Credentials obtained"
    Write-Output "Connecting to O365"
    Connect-MsolService -Credential $credential
}
catch{
    Write-Output "Connecttion to O365 failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
    $Out = "Time Elapsed = $($Timer.Elapsed.ToString())
    Connecttion to O365 failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
    $Timer.Stop()
    Out-File -Encoding Ascii -FilePath $res -inputObject $Out
    Return
}
try{
    Write-Output "Connected to O365"
    Write-Output "Conneting to Azure AD"
    Connect-AzureAD -Credential $credential
}
catch{
    Write-Output "Connecttion to Azure AD failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
    $Out = "Time Elapsed = $($Timer.Elapsed.ToString())
    Connecttion to Azure AD failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
    $Timer.Stop()
    Out-File -Encoding Ascii -FilePath $res -inputObject $Out
    Return
}
$g = new-object Microsoft.Open.AzureAD.Model.GroupIdsForMembershipCheck
$g.GroupIds = "a5f37d5e-5f32-4779-a710-51e4342ffd29"
try{
    Write-Output "Connected to Azure AD"
    Write-Output "Get all Users"
    $Users = Get-MsolUser -EnabledFilter EnabledOnly -All
}
catch{
    Write-Output "Failed getting users"
    $out = "Time Elapsed = $($Timer.Elapsed.ToString())
    $($_.Invocationinfo.MyCommand) at position $($_.Invocationinfo.positionmessage) failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
    $Timer.Stop()
    Out-File -Encoding Ascii -FilePath $res -inputObject $Out
    Return
}
If ($Users){
    $out = @()
    foreach($User in $Users){
        Write-Output "Processing user $($User)"
        if($User.ImmutableId){
            $AccountType = "On-Prem"
        } Else {
            $AccountType = "AzureAD"
        }
        If($User.StrongAuthenticationRequirements){
            $UserMFASetting = $User.StrongAuthenticationRequirements.state
        } Else {
            $UserMFASetting = "Disabled"
        }
        If (Select-AzureADGroupIdsUserIsMemberOf -ObjectId $User.ObjectId -GroupIdsForMembershipCheck $g){
            $MFAConditionalAccess = "Enabled"
        } Else {
            $MFAConditionalAccess = "Disabled"
        }
        $O = New-Object psobject -Property @{
            UserPrincipalName = $User.UserPrincipalName
            Type = $AccountType
            UserMFASetting = $UserMFASetting
            MFAConditionalAccess = $MFAConditionalAccess
        }
        $out += $O
    }
} Else {
    Write-Output "0 Users retrieved from O365"
    $Out = "Time Elapsed = $($Timer.Elapsed.ToString())
    0 Users retrieved from O365"
    Out-File -Encoding Ascii -FilePath $res -inputObject $Out
    Return
}
$Out += $Timer.Elapsed.ToString()
Write-Output "Completed execution Time elapsed: $($Timer.Elapsed.ToString())"
$Timer.Stop()
Out-File -Encoding Ascii -FilePath $res -inputObject $out
Return