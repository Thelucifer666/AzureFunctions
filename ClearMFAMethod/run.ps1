$requestBody = Get-Content $req -Raw | ConvertFrom-Json
$UPN = $requestBody.UserPrincipalName
$Result = "Failure"
$AccountType = $null
$UserMFASetting = $null
$MFAPolicy = $null
$DefaultMFAMethod = $null
$ClearMFAMethod = "False"
$ForceTokenExpiry = "False"
Write-Output "$($UPN)"
Write-Output "Function started execution at $(Get-Date)"
If ($UPN){
    If($UPN.Count -eq 1){
        Write-Output "Gathering credentials"
        $Timer =  [system.diagnostics.stopwatch]::StartNew()
        try{
            $username = $Env:user
            $pw = $Env:password
            $keypath = "D:\home\site\wwwroot\bin\keys\PassEncryptKey.key"
            $secpassword = $pw | ConvertTo-SecureString -Key (Get-Content $keypath)
            $credential = New-Object System.Management.Automation.PSCredential ($username, $secpassword)
        }
        catch{
            $O = New-Object PSCustomObject -Property @{
                Result = $Result
                UserPrincipalName = $UPN
                Type = $AccountType
                MFAPolicy = $MFAPolicy
                UserMFASetting = $UserMFASetting
                DefaultMFAMethod = $DefaultMFAMethod
                ClearMFAMethod = $ClearMFAMethod
                ForceTokenExpiry = $ForceTokenExpiry
                Error = "Creating the credential object failed.
                $($_.Invocationinfo.MyCommand) at position $($_.Invocationinfo.positionmessage) failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            }
            $Out = $O | ConvertTo-Json
            $Timer.Stop()
            Out-File -Encoding Ascii -FilePath $res -inputObject $Out
            Return
        }
        try{
            Write-Output "Credentials obtained"
            Write-Output "Updating PS Module Path environment variable"
            $env:PSModulePath = $env:PSModulePath + ";d:\home\site\wwwroot\bin\modules\"
            Write-Output "Importing MSOnline PS Module"
            Import-Module MSOnline -ErrorAction Stop
            Write-Output "Imported MSOnline PS Module"
            Write-Output "Importing AzureAD PS Module"
            Import-Module AzureAD -ErrorAction Stop
            Write-Output "Imported AzureAD PS Module"
        }
        catch{
            $O = New-Object PSCustomObject -Property @{
                Result = $Result
                UserPrincipalName = $UPN
                Type = $AccountType
                MFAPolicy = $MFAPolicy
                UserMFASetting = $UserMFASetting
                DefaultMFAMethod = $DefaultMFAMethod
                ClearMFAMethod = $ClearMFAMethod
                ForceTokenExpiry = $ForceTokenExpiry
                Error = "Importing module failed.
                $($_.Invocationinfo.MyCommand) at position $($_.Invocationinfo.positionmessage) failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            }
            $Out = $O | ConvertTo-Json
            $Timer.Stop()
            Out-File -Encoding Ascii -FilePath $res -inputObject $Out
            Return
        }
        try{
            Write-Output "Connecting to O365"
            Connect-MsolService -Credential $credential -ErrorAction Stop
            Write-Output "Connected to O365"
        }
        catch{
            Write-Output "Connecttion to O365 failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            $O = New-Object PSCustomObject -Property @{
                Result = $Result
                UserPrincipalName = $UPN
                Type = $AccountType
                MFAPolicy = $MFAPolicy
                UserMFASetting = $UserMFASetting
                DefaultMFAMethod = $DefaultMFAMethod
                ClearMFAMethod = $ClearMFAMethod
                ForceTokenExpiry = $ForceTokenExpiry
                Error = "Connecttion to O365 failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            }
            $Out = $O | ConvertTo-Json
            $Timer.Stop()
            Out-File -Encoding Ascii -FilePath $res -inputObject $Out
            Return
        }
        try{
            Write-Output "Conneting to Azure AD"
            Connect-AzureAD -Credential $credential -ErrorAction Stop
            Write-Output "Connected to Azure AD"
        }
        catch{
            Write-Output "Connecttion to Azure AD failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            $O = New-Object PSCustomObject -Property @{
                Result = $Result
                UserPrincipalName = $UPN
                Type = $AccountType
                MFAPolicy = $MFAPolicy
                UserMFASetting = $UserMFASetting
                DefaultMFAMethod = $DefaultMFAMethod
                ClearMFAMethod = $ClearMFAMethod
                ForceTokenExpiry = $ForceTokenExpiry
                Error = "Connecttion to Azure AD failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            }
            $Out = $O | ConvertTo-Json
            $Timer.Stop()
            Out-File -Encoding Ascii -FilePath $res -inputObject $Out
            Return
        }
        try{
            Write-Output "Get user properties from MSOnine"
            $User = Get-MsolUser -UserPrincipalName $UPN -ErrorAction SilentlyContinue
            Write-Output "$User"
            $g = New-Object Microsoft.Open.AzureAD.Model.GroupIdsForMembershipCheck
            $g.GroupIds = "a5f37d5e-5f32-4779-a710-51e4342ffd29","0b07dff6-392e-438c-9298-20c1a6a86077","4e086602-3259-40e2-b7c7-92cc7d99dded"
            if ($User){
                Write-Output "Processing user $($User.UserPrincipalName)"
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
                If (Select-AzureADGroupIdsUserIsMemberOf -ObjectId $User.ObjectId -GroupIdsForMembershipCheck $g -ErrorAction SilentlyContinue){
                    $MFAPolicy = "Enabled"
                    $MFAg = New-Object Microsoft.Open.AzureAD.Model.GroupIdsForMembershipCheck
                    $MFAg.GroupIds = "a5f37d5e-5f32-4779-a710-51e4342ffd29"
                    If (Select-AzureADGroupIdsUserIsMemberOf -ObjectId $User.ObjectId -GroupIdsForMembershipCheck $MFAg -ErrorAction SilentlyContinue){
                        $MFAPolicy = "Enabled"
                    } Else {
                        Add-AzureADGroupMember -ObjectId "a5f37d5e-5f32-4779-a710-51e4342ffd29" -RefObjectId $User.ObjectId -ErrorAction Stop
                        $MFAPolicy = "Enabled"
                    }
                } Else {
                    $MFAPolicy = "Disabled"
                    Add-AzureADGroupMember -ObjectId "a5f37d5e-5f32-4779-a710-51e4342ffd29" -RefObjectId $User.ObjectId -ErrorAction Stop
                    $MFAPolicy = "Enabled"
                }
                If($User.StrongAuthenticationMethods){
                    $DefaultMFAMethod = ($User.StrongAuthenticationMethods | Where-Object{$_.IsDefault-eq "True"}).MethodType
                    $stm = @()
                    Set-MsolUser -ObjectId $User.ObjectId -StrongAuthenticationMethods $stm -ErrorAction Stop
                    $DefaultMFAMethod = "Not Set"
                    $ClearMFAMethod = "True"
                } Else {
                    $DefaultMFAMethod = "Not Set"
                }
                Revoke-AzureADUserAllRefreshToken -ObjectId $User.ObjectId -ErrorAction Stop
                $ForceTokenExpiry = "True"
                $Result = "Success"
                
                $O = New-Object psobject -Property @{
                    Result = $Result
                    UserPrincipalName = $User.UserPrincipalName
                    Type = $AccountType
                    MFAPolicy = $MFAPolicy
                    UserMFASetting = $UserMFASetting
                    DefaultMFAMethod = $DefaultMFAMethod
                    ClearMFAMethod = $ClearMFAMethod
                    ForceTokenExpiry = $ForceTokenExpiry
                    Error = $null
                }
                $Out = $O | ConvertTo-Json
                $Timer.Stop()
                Out-File -Encoding Ascii -FilePath $res -inputObject $out
                Return
            } Else {
                $O = New-Object psobject -Property @{
                    Result = $Result
                    UserPrincipalName = $UPN
                    Type = $AccountType
                    MFAPolicy = $MFAPolicy
                    UserMFASetting = $UserMFASetting
                    DefaultMFAMethod = $DefaultMFAMethod
                    ClearMFAMethod = $ClearMFAMethod
                    ForceTokenExpiry = $ForceTokenExpiry
                    Error = "Please provide a valid UserPrincipalName"
                }
                $Out = $O | ConvertTo-Json
                $Timer.Stop()
                Out-File -Encoding Ascii -FilePath $res -inputObject $Out
                Return
            }
        } catch {
            $O = New-Object PSCustomObject -Property @{
                Result = $Result
                UserPrincipalName = $UPN
                Type = $AccountType
                MFAPolicy = $MFAPolicy
                UserMFASetting = $UserMFASetting
                DefaultMFAMethod = $DefaultMFAMethod
                ClearMFAMethod = $ClearMFAMethod
                ForceTokenExpiry = $ForceTokenExpiry
                Error = "Failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            }
            $Out = $O | ConvertTo-Json
            $Timer.Stop()
            Out-File -Encoding Ascii -FilePath $res -inputObject $Out
            Return
        }
    } Else {
        $O = New-Object psobject -Property @{
            Result = $Result
            UserPrincipalName = $UPN
            Type = $AccountType
            MFAPolicy = $MFAPolicy
            UserMFASetting = $UserMFASetting
            DefaultMFAMethod = $DefaultMFAMethod
            ClearMFAMethod = $ClearMFAMethod
            ForceTokenExpiry = $ForceTokenExpiry
            Error = "Please provide ONE valid UserPrincipalName"
        }
        $Out = $O | ConvertTo-Json
        $Timer.Stop()
        Out-File -Encoding Ascii -FilePath $res -inputObject $Out
        Return
    }
} Else {
    $O = New-Object psobject -Property @{
        Result = $Result
        UserPrincipalName = $UPN
        Type = $AccountType
        MFAPolicy = $MFAPolicy
        UserMFASetting = $UserMFASetting
        DefaultMFAMethod = $DefaultMFAMethod
        ClearMFAMethod = $ClearMFAMethod
        ForceTokenExpiry = $ForceTokenExpiry
        Error = "Please provide a valid UserPrincipalName"
    }
    $Out = $O | ConvertTo-Json
    $Timer.Stop()
    Out-File -Encoding Ascii -FilePath $res -inputObject $Out
    Return
}