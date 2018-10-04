# SetMFA - PowerShell

This `HttpTrigger` enables or disables Multi-Factor Authentication for a given user

## How it works

The function accepts a User's `UserPrincipalName`, `Action` (`Enable`/`Disable`), `Force` in JSON  format, verifies if this is a valid user account. For a valid user account, this function will enable or disable MFA based on the input and output the status in JSON format.

## Learn more

You can learn more on Azure Functions [here](https://docs.microsoft.com/en-us/azure/azure-functions/)