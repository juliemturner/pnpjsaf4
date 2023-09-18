# PnPjs - Azure Function v4 Support

This sample project demonstrates how to use the PnPjs SDK inside of an Azure Function v4 with Application Insights, one of the most common scenarios when building extensibility solutions for Microsoft 365.

## AzureIdentity

AzureIdentity is an SDK that provides Azure Active Directory (Azure AD) token authentication through a set of convenient TokenCredential implementations. [More information](https://github.com/Azure/azure-sdk-for-js/blob/main/sdk/identity/identity/README.md)

This sample implements PnPjs security expecting Managed Identity has been configured for the Azure Function. When in local development mode you can log into Azure by including the Azure Account extension for Visual Studio Code. See the more information link above for other supported authentication scenarios like Client Secret and Certificate.

## Application Insights

[more infomration](https://github.com/microsoft/ApplicationInsights-node.js#readme)