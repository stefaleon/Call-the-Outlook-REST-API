## [Write an ASP.NET MVC Web app to get Outlook mail, calendar, and contacts](https://docs.microsoft.com/en-us/outlook/rest/dotnet-tutorial)

    02/20/2018
    20 minutes to read
    Contributors
        Jason Johnston Florian Geiger SiavasFiroozbakht 

The purpose of this guide is to walk through the process of creating a simple ASP.NET MVC C# app that retrieves messages in Office 365 or Outlook.com. The source code in this repository is what you should end up with if you follow the steps outlined here.

This tutorial will use the Microsoft Authentication Library (MSAL) to make OAuth2 calls and the Microsoft Graph Client Library to call the Mail API. Microsoft recommends using the Microsoft Graph to access Outlook mail, calendar, and contacts. You should use the Outlook APIs directly (via https://outlook.office.com/api) only if you require a feature that is not available on the Graph endpoints. 


### Create the app

Let's dive right in! In Visual Studio, create a new Visual C# ASP.NET Web Application using .NET Framework 4.5. Name the application dotnet-tutorial.

Select the MVC template. Click the Change Authentication button and choose "No Authentication".