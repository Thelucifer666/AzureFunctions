# TestMFA - PowerShell

This `HttpTrigger` verifies the user's individual MFA Multi-Factor Authentication setting from O365 as well as check if the user is part of the Azure/AD groups assigned for Conditional Access Polocies for MFA.

## How it works

The function accepts a User's `UserPrincipalName` in JSON  format, verifies if this is a valid user account. For a valid user account, this function will query O365 for user's MFA setting and also query AzureAD for group memberships and output the status of the each in JSON format.

## Learn more

You can learn more on Azure Functions [here](https://docs.microsoft.com/en-us/azure/azure-functions/)