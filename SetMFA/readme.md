# SetMFA - PowerShell

This `HttpTrigger` enables or disables Multi-Factor Authentication for a given user

## How it works

The function accepts a User's `UserPrincipalName`, `Action` (`Enable`/`Disable`), `Force` in JSON  format, verifies if this is a valid user account. For a valid user account, this function will enable or disable MFA by adding/removing user to the conditional access policy group and output the status in JSON format. This function will also revoke refresh token for the user so the policies can be applied at the earliest.

## Learn more

You can learn more on Azure Functions [here](https://docs.microsoft.com/en-us/azure/azure-functions/)