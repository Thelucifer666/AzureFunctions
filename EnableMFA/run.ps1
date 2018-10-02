# POST method: $req
$in = [PSCustomObject]@{
    UserPrincipalName = "ravi","Dan"
}
$req = "C:\In.txt"
ConvertTo-Json -InputObject $in | Out-File $req

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
            $Out = New-Object PSCustomObject -Property @{
                Result = "Failure"
                Error = "Creating the credential object failed.
                $($_.Invocationinfo.MyCommand) at position $($_.Invocationinfo.positionmessage) failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            }
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
            $Out = New-Object PSCustomObject -Property @{
                Result = "Failure"
                Error = "Importing module failed.
                $($_.Invocationinfo.MyCommand) at position $($_.Invocationinfo.positionmessage) failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            }
            $Timer.Stop()
            Out-File -Encoding Ascii -FilePath $res -inputObject $Out
            Return
        }
        # Connect to MSOnline
        try{
            Write-Output "Connecting to O365"
            Connect-MsolService -Credential $credential
        }
        catch{
            Write-Output "Connecttion to O365 failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            $Out = New-Object PSCustomObject -Property @{
                Result = "Failure"
                Error = "Connecttion to O365 failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            }
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
            $Out = New-Object PSCustomObject -Property @{
                Result = "Failure"
                Error = "Connecttion to Azure AD failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            }
            $Timer.Stop()
            Out-File -Encoding Ascii -FilePath $res -inputObject $Out
            Return
        }
        try{
            Write-Output "Get user properties from MSOnine"
            $User = Get-MsolUser -UserPrincipalName $UPN
            $g = new-object Microsoft.Open.AzureAD.Model.GroupIdsForMembershipCheck
            $g.GroupIds = "a5f37d5e-5f32-4779-a710-51e4342ffd29","0b07dff6-392e-438c-9298-20c1a6a86077","4e086602-3259-40e2-b7c7-92cc7d99dded"

        } catch {
            $Out = New-Object PSCustomObject -Property @{
                Result = "Failure"
                Error = "Failed with the following exception message: $($_.Exception.Message); error code: $($_.Exception.ErrorCode); Inner exception: $($_.Exception.InnerException); HResult: $($_.Exception.HResult); Category: $($_.CategoryInfo.Category)"
            }
        }
    } Else {
        $Out = New-Object psobject -Property @{
            Result = "Failure"
            Error = "Please provide one valid UserPrincipaName"
        }
    }
} Else {
    $Out = New-Object psobject -Property @{
        Result = "Failure"
        Error = "Please provide a valid UserPrincipaName"
    }
    Out-File -Encoding Ascii -FilePath $res -inputObject $Out
}
Out-File -Encoding Ascii -FilePath $res -inputObject $Out