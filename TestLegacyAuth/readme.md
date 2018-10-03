# TestLegacyAuth - PowerShell

This `HttpTrigger` verifies the user is part of the Azure/AD groups assigned for Conditional Access Polocies for disabling Legacy Authentication protocols.

## How it works

The function accepts a User's `UserPrincipalName` in JSON  format, verifies if this is a valid user account. For a valid user account, this function queries AzureAD for user's group memberships to check the if the user is member of the groups assigned to the conditional access policy for disabling legacy authentication and outputs the status in JSON format.

## Learn more

You can learn more on Azure Functions [here](https://docs.microsoft.com/en-us/azure/azure-functions/)