# TestMAM - PowerShell

This `HttpTrigger` verifies if the user has been added groups assigned to Mobile Application Management (MAM) conditional access policy in Azure AD and check if the necessary licenses are assigned.

## How it works

The function accepts a User's `UserPrincipalName` in JSON  format, verifies if this is a valid user account. For a valid user account, this function will query Azure AD for group memberships for conditional access policies for MAM and also check if the necessary Enterprise Mobility and Security (EMS) license has been assigned to the user for the polocies to take effect. output the status in JSON format. 

## Learn more

You can learn more on Azure Functions [here](https://docs.microsoft.com/en-us/azure/azure-functions/)