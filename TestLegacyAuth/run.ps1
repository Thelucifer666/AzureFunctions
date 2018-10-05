$requestBody = Get-Content $req -Raw | ConvertFrom-Json
$UPN = $requestBody.UserPrincipalName
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
                Result = "Failure"
                UserPrincipalName = $UPN
                Type = $null
                LegacyAuthPolicy = $null
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
            Write-Output "Importing AzureAD PS Module"
            Import-Module AzureAD -ErrorAction Stop
            Write-Output "Imported AzureAD PS Module"
        }
        catch{
            $O = New-Object PSCustomObject -Property @{
                Result = "Failure"
                UserPrincipalName = $UPN
                Type = $null
                LegacyAuthPolicy = $null
                Error = "Importing module failed.
                $($_.Invocationinfo.MyCommand) at position $($_.Invocationinfo.positionmessage) failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
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
                Result = "Failure"
                UserPrincipalName = $UPN
                Type = $null
                LegacyAuthPolicy = $null
                Error = "Connecttion to Azure AD failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            }
            $Out = $O | ConvertTo-Json
            $Timer.Stop()
            Out-File -Encoding Ascii -FilePath $res -inputObject $Out
            Return
        }
        try{
            Write-Output "Get user properties from AzureAD"
            $User = Get-AzureADUser -Filter "UserPrincipalName eq '$($UPN)'" -ErrorAction SilentlyContinue
            Write-Output "$User"
            $g = new-object Microsoft.Open.AzureAD.Model.GroupIdsForMembershipCheck
            $g.GroupIds = "f226447b-9481-497e-b171-996a77f8e81c", "0b07dff6-392e-438c-9298-20c1a6a86077","4e086602-3259-40e2-b7c7-92cc7d99dded"
            if ($User){
                Write-Output "Processing user $($User.UserPrincipalName)"
                if($User.ImmutableId){
                    $AccountType = "On-Prem"
                } Else {
                    $AccountType = "AzureAD"
                }
                If (Select-AzureADGroupIdsUserIsMemberOf -ObjectId $User.ObjectId -GroupIdsForMembershipCheck $g -ErrorAction SilentlyContinue){
                    $LegacyAuthPolicy = "Enabled"
                } Else {
                    $LegacyAuthPolicy = "Disabled"
                }
                $O = New-Object psobject -Property @{
                    Result = "Success"
                    UserPrincipalName = $User.UserPrincipalName
                    Type = $AccountType
                    LegacyAuthPolicy = $LegacyAuthPolicy
                    Error = $null
                }
                $Out = $O | ConvertTo-Json
                $Timer.Stop()
                Out-File -Encoding Ascii -FilePath $res -inputObject $out
                Return
            } Else {
                $O = New-Object psobject -Property @{
                    Result = "Failure"
                    UserPrincipalName = $UPN
                    Type = $null
                    LegacyAuthPolicy = $null
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
                LegacyAuthPolicy = $null
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
            LegacyAuthPolicy = $null
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
        LegacyAuthPolicy = $null
        Error = "Please provide a valid UserPrincipalName"
    }
    $Out = $O | ConvertTo-Json
    $Timer.Stop()
    Out-File -Encoding Ascii -FilePath $res -inputObject $Out
    Return
}