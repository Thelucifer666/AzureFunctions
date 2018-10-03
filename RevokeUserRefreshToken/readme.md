# RevokeUserRefreshToken - PowerShell

This `HttpTrigger` invalidates the refresh tokens issued to applications and/or session cookies in a browser for a user.

## How it works

The function accepts a User's `UserPrincipalName` in JSON  format, verifies if this is a valid user account. For a valid user account, this function will invalidate the refresh tokens issued to applications and/or session cookies in a browser.

## Learn more

You can learn more on Azure Functions [here](https://docs.microsoft.com/en-us/azure/azure-functions/)