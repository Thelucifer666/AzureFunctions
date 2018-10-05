# SetLegacyAuth - PowerShell

This `HttpTrigger` enables or disables Legacy Authentication for a given user

## How it works

The function accepts a User's `UserPrincipalName`, `Action` (`Enable`/`Disable`) in JSON  format, verifies if this is a valid user account. For a valid user account, this function will add or remove a user from the groups assigned for Conditional access policy for disabling legacy authent1ication and output the status in JSON format. This function will also revoke refresh token for the user so the policies can be applied at the earliest.

## Learn more

You can learn more on Azure Functions [here](https://docs.microsoft.com/en-us/azure/azure-functions/)