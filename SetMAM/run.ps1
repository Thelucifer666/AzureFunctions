$requestBody = Get-Content $req -Raw | ConvertFrom-Json
$UPN = $requestBody.UserPrincipalName
$Result = "Failure"
$AccountType = $null
$MAMPolicy = $null
$EMSLicenseStatus = $null
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
                MAMPolicy = $MAMPolicy
                EMSLicenseStatus = $EMSLicenseStatus
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
            Write-Output "Importing AzureAD PS Module"
            Import-Module AzureAD -ErrorAction Stop
            Write-Output "Imported AzureAD PS Module"
        }
        catch{
            $O = New-Object PSCustomObject -Property @{
                Result = $Result
                UserPrincipalName = $UPN
                Type = $AccountType
                MAMPolicy = $MAMPolicy
                EMSLicenseStatus = $EMSLicenseStatus
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
                MAMPolicy = $MAMPolicy
                EMSLicenseStatus = $EMSLicenseStatus
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
            $User = Get-AzureADUser -Filter "Userprincipalname eq '$($UPN)'" -ErrorAction SilentlyContinue
            Write-Output "$User"
            $g = New-Object Microsoft.Open.AzureAD.Model.GroupIdsForMembershipCheck
            $g.GroupIds = "0b07dff6-392e-438c-9298-20c1a6a86077","4e086602-3259-40e2-b7c7-92cc7d99dded","daad0aa3-77af-43ae-bd44-31961f4cbe2a","e96bf8a2-c509-4bd1-acca-6889bdea352f"
            if ($User){
                Write-Output "Processing user $($User.UserPrincipalName)"
                if($User.ImmutableId){
                    $AccountType = "On-Prem"
                } Else {
                    $AccountType = "AzureAD"
                }
                If($requestBody.Action -eq "Enable"){
                    If (Select-AzureADGroupIdsUserIsMemberOf -ObjectId $User.ObjectId -GroupIdsForMembershipCheck $g -ErrorAction SilentlyContinue){
                        $MAMPolicy = "Enabled"
                        $MAMg = New-Object Microsoft.Open.AzureAD.Model.GroupIdsForMembershipCheck
                        $MAMg.GroupIds = "daad0aa3-77af-43ae-bd44-31961f4cbe2a"
                        If (Select-AzureADGroupIdsUserIsMemberOf -ObjectId $User.ObjectId -GroupIdsForMembershipCheck $MAMg -ErrorAction SilentlyContinue){
                            $MAMPolicy = "Enabled"
                        } Else {
                            Add-AzureADGroupMember -ObjectId "daad0aa3-77af-43ae-bd44-31961f4cbe2a" -RefObjectId $User.ObjectId -ErrorAction Stop
                            $MAMPolicy = "Enabled"
                        }
                    } Else {
                        $MAMPolicy = "Disabled"
                        Add-AzureADGroupMember -ObjectId "daad0aa3-77af-43ae-bd44-31961f4cbe2a" -RefObjectId $User.ObjectId -ErrorAction Stop
                        $MAMPolicy = "Enabled"
                    }
                    Revoke-AzureADUserAllRefreshToken -ObjectId $User.ObjectId -ErrorAction Stop
                    $ForceTokenExpiry = "True"
                    Write-Output "Get the Enterprise Mobility and Security License SKUId"
                    $EMSSkuId = (Get-AzureADSubscribedSku | Where-Object {$_.SkuPartNumber -match 'EMS'}).SkuId
                    Write-Output "Verify if User has EMS License assigned"
                    $UserEMSLicense = $User.AssignedLicenses | Where-Object {$_.SkuID -eq $EMSSkuId}
                    If($UserEMSLicense){
                        $EMSLicenseStatus = "Assigned"
                    } Else {
                        $EMSLicenseStatus = "Not Assigned"
                    }
                    $Result = "Success"
                } elseif ($requestBody.Action -eq "Disable") {
                    If (Select-AzureADGroupIdsUserIsMemberOf -ObjectId $User.ObjectId -GroupIdsForMembershipCheck $g -ErrorAction SilentlyContinue){
                        $MAMPolicy = "Enabled"
                        foreach($group in (Select-AzureADGroupIdsUserIsMemberOf -ObjectId $User.ObjectId -GroupIdsForMembershipCheck $g -ErrorAction SilentlyContinue)){
                            Remove-AzureADGroupMember -ObjectId $group -MemberId $User.ObjectId -ErrorAction Stop
                        }
                        $MAMPolicy = "Disabled"
                    } Else {
                        $MAMPolicy = "Disabled"
                    }
                    Revoke-AzureADUserAllRefreshToken -ObjectId $User.ObjectId -ErrorAction Stop
                    $ForceTokenExpiry = "True"
                    Write-Output "Get the Enterprise Mobility and Security License SKUId"
                    $EMSSkuId = (Get-AzureADSubscribedSku | Where-Object {$_.SkuPartNumber -match 'EMS'}).SkuId
                    Write-Output "Verify if User has EMS License assigned"
                    $UserEMSLicense = $User.AssignedLicenses | Where-Object {$_.SkuID -eq $EMSSkuId}
                    If($UserEMSLicense){
                        $EMSLicenseStatus = "Assigned"
                    } Else {
                        $EMSLicenseStatus = "Not Assigned"
                    }
                    $Result = "Success"
                }
                $O = New-Object psobject -Property @{
                    Result = $Result
                    UserPrincipalName = $User.UserPrincipalName
                    Type = $AccountType
                    MAMPolicy = $MAMPolicy
                    EMSLicenseStatus = $EMSLicenseStatus
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
                    MAMPolicy = $MAMPolicy
                    EMSLicenseStatus = $EMSLicenseStatus
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
                MAMPolicy = $MAMPolicy
                EMSLicenseStatus = $EMSLicenseStatus
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
            MAMPolicy = $MAMPolicy
            EMSLicenseStatus = $EMSLicenseStatus
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
        MAMPolicy = $MAMPolicy
        EMSLicenseStatus = $EMSLicenseStatus
        ForceTokenExpiry = $ForceTokenExpiry
        Error = "Please provide a valid UserPrincipalName"
    }
    $Out = $O | ConvertTo-Json
    $Timer.Stop()
    Out-File -Encoding Ascii -FilePath $res -inputObject $Out
    Return
}