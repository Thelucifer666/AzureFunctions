$requestBody = Get-Content $req -Raw | ConvertFrom-Json
$UPN = $requestBody.UserPrincipalName
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
                Result = "Failure"
                UserPrincipalName = $UPN
                Type = $null
                MFAConditionalAccess = $null
                UserMFASetting = $null
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
            Import-Module MSOnline
            Write-Output "Imported MSOnline PS Module"
            Write-Output "Importing AzureAD PS Module"
            Import-Module AzureAD
            Write-Output "Imported MSOnline PS Module"
        }
        catch{
            $O = New-Object PSCustomObject -Property @{
                Result = "Failure"
                UserPrincipalName = $UPN
                Type = $null
                MFAConditionalAccess = $null
                UserMFASetting = $null
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
            Connect-MsolService -Credential $credential
        }
        catch{
            Write-Output "Connecttion to O365 failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            $O = New-Object PSCustomObject -Property @{
                Result = "Failure"
                UserPrincipalName = $UPN
                Type = $null
                MFAConditionalAccess = $null
                UserMFASetting = $null
                Error = "Connecttion to O365 failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            }
            $Out = $O | ConvertTo-Json
            $Timer.Stop()
            Out-File -Encoding Ascii -FilePath $res -inputObject $Out
            Return
        }
        try{
            Write-Output "Connected to O365"
            Write-Output "Conneting to Azure AD"
            Connect-AzureAD -Credential $credential
            Write-Output "Connected to Azure AD"
        }
        catch{
            Write-Output "Connecttion to Azure AD failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            $O = New-Object PSCustomObject -Property @{
                Result = "Failure"
                UserPrincipalName = $UPN
                Type = $null
                MFAConditionalAccess = $null
                UserMFASetting = $null
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
            $g = new-object Microsoft.Open.AzureAD.Model.GroupIdsForMembershipCheck
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
                    $MFAConditionalAccess = "Enabled"
                } Else {
                    $MFAConditionalAccess = "Disabled"
                }
                $O = New-Object psobject -Property @{
                    Result = "Success"
                    UserPrincipalName = $User.UserPrincipalName
                    Type = $AccountType
                    UserMFASetting = $UserMFASetting
                    MFAConditionalAccess = $MFAConditionalAccess
                    Error = $null
                }
                $Out = $O | ConvertTo-Json
                $Timer.Stop()
                Out-File -Encoding Ascii -FilePath $res -inputObject $out
            } Else {
                $O = New-Object psobject -Property @{
                    Result = "Failure"
                    UserPrincipalName = $UPN
                    Type = $null
                    MFAConditionalAccess = $null
                    UserMFASetting = $null
                    Error = "Please provide a valid UserPrincipalName"
                }
                $Out = $O | ConvertTo-Json
                $Timer.Stop()
                Out-File -Encoding Ascii -FilePath $res -inputObject $Out
                Return
            }
        } catch {
            $O = New-Object PSCustomObject -Property @{
                Result = "Failure"
                UserPrincipalName = $UPN
                Type = $null
                MFAConditionalAccess = $null
                UserMFASetting = $null
                Error = "Failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            }
            $Out = $O | ConvertTo-Json
            $Timer.Stop()
            Out-File -Encoding Ascii -FilePath $res -inputObject $Out
            Return
        }
    } Else {
        $O = New-Object psobject -Property @{
            Result = "Failure"
            UserPrincipalName = $UPN
            Type = $null
            MFAConditionalAccess = $null
            UserMFASetting = $null
            Error = "Please provide ONE valid UserPrincipalName"
        }
        $Out = $O | ConvertTo-Json
        $Timer.Stop()
        Out-File -Encoding Ascii -FilePath $res -inputObject $Out
        Return
    }
} Else {
    $O = New-Object psobject -Property @{
        Result = "Failure"
        UserPrincipalName = $UPN
        Type = $null
        MFAConditionalAccess = $null
        UserMFASetting = $null
        Error = "Please provide a valid UserPrincipalName"
    }
    $Out = $O | ConvertTo-Json
    $Timer.Stop()
    Out-File -Encoding Ascii -FilePath $res -inputObject $Out
    Return
}