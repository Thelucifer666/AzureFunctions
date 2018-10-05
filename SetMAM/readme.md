# SetMAM - PowerShell

This `HttpTrigger` enables or disables Mobile Application Management for a given user

## How it works

The function accepts a User's `UserPrincipalName`, `Action` (`Enable`/`Disable`) in JSON  format, verifies if this is a valid user account. For a valid user account, this function will enable or disable MAM by adding or removing the user to/from the group assigned to the conditional access policy and output the status in JSON format. This function will also revoke refresh token for the user so the policies can be applied at the earliest.

## Learn more

You can learn more on Azure Functions [here](https://docs.microsoft.com/en-us/azure/azure-functions/)