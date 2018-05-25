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


### Designing the app

Our app will be very simple. When a user visits the site, they will see a button to log in and view their email. Clicking that button will take them to the Azure login page where they can login with their Office 365 or Outlook.com account and grant access to our app. Finally, they will be redirected back to our app, which will display a list of the most recent email in the user's inbox.

Let's begin by replacing the stock home page with a simpler one. Open the ./Views/Home/Index.cshtml file. Replace the existing code with the following code.


Contents of the ./Views/Home/Index.cshtml file

```
@{
    ViewBag.Title = "Home Page";
}

<div class="jumbotron">
    <h1>ASP.NET MVC Tutorial</h1>
    <p class="lead">This sample app uses the Mail API to read messages in your inbox.</p>
    <p><a href="#" class="btn btn-primary btn-lg">Click here to login</a></p>
</div>
```


This is basically repurposing the jumbotron element from the stock home page, and removing all of the other elements. The button doesn't do anything yet, but the home page should now look like the following.



Let's also modify the stock error page so that we can pass an error message and display it. Replace the contents of ./Views/Shared/Error.cshtml with the following code.



Contents of the ./Views/Shared/Error.cshtml file

```
<!DOCTYPE html>
<html>
<head>
    <meta name="viewport" content="width=device-width" />
    <title>Error</title>
</head>
<body>
    <hgroup>
        <h1>Error.</h1>
        <h2>An error occurred while processing your request.</h2>
    </hgroup>
    <div class="alert alert-danger">@ViewBag.Message</div>
    @if (!string.IsNullOrEmpty(ViewBag.Debug))
    {
    <pre><code>@ViewBag.Debug</code></pre>
    }
</body>
</html>
```


Finally add an action to the HomeController class to invoke the error view.

The Error action in ./Controllers/HomeController.cs

```
public ActionResult Error(string message, string debug)
{
    ViewBag.Message = message;
    ViewBag.Debug = debug;
    return View("Error");
}
```



## Register the app

! Important

New app registrations should be created and managed in the new Application Registration Portal to be compatible with Outlook.com. 

Account requirements

In order to use the Application Registration Portal, you need either an Office 365 work or school account, or a Microsoft account. 

REST API availability

The REST API is currently enabled on all Office 365 accounts that have Exchange Online, and all Outlook.com accounts.


Head over to the [Application Registration Portal](https://apps.dev.microsoft.com) to quickly get an application ID and secret.

1. Using the Sign in link, sign in with either your Microsoft account (Outlook.com), or your work or school account (Office 365).

2. Click the Add an app button. Enter `CallOutlookApi` for the name and click Create application.

3. Locate the Application Secrets section, and click the Generate New Password button. Copy the password now and save it to a safe place. Once you've copied the password, click Ok.

4. Locate the Platforms section, and click Add Platform. Choose Web, then enter `http://localhost:<PORT>`, replacing `<PORT>` with the port number your project is using under Redirect URLs.
    
    Tip

    You can find the port number your project is using by selecting the project in Solution Explorer, then checking the value of URL under Development Server in the Properties window.

5. Click Save to complete the registration. Copy the Application Id and save it along with the password you copied earlier. We'll need those values soon.



## Implementing OAuth2

Our goal in this section is to make the link on our home page initiate the [OAuth2 Authorization Code Grant flow with Azure AD](https://msdn.microsoft.com/en-us/library/azure/dn645542.aspx). To make things easier, we'll use the [Microsoft Authentication Library (MSAL)](https://www.nuget.org/packages/Microsoft.Identity.Client) to handle our OAuth requests.

First, let's create a separate config file to hold the OAuth settings for the app. Right-click the solution in Solution Explorer, choose Add, then New Item.... Select Web Configuration File, then enter AzureOauth.config in the Name field. Click Add.

Open the AzureOauth.config file and replace its entire contents with the following.

XML

```
<appSettings>
    <add key="ida:AppID" value="YOUR APP ID" />
    <add key="ida:AppPassword" value="YOUR APP PASSWORD" />
    <add key="ida:RedirectUri" value="http://localhost:10800" />
    <add key="ida:AppScopes" value="User.Read Mail.Read" />
</appSettings>
```
Replace the value of the ida:AppID key with the application ID you generated above, and replace the value of the ida:AppPassword key with the password you generated above. If the value of your redirect URI is different, be sure to update the value of ida:RedirectUri.

Now open the Web.config file. Find the line with the <appSettings> element, and change it to the following.

XML

```
<appSettings file="AzureOauth.config">
```

This will cause ASP.NET to add the keys from the AzureOauth.config file at runtime. By keeping these values in a separate file, we make it less likely that we'll accidentally commit them to source control.

The next step is to install the OWIN middleware, MSAL, and Graph libraries from NuGet. On the Visual Studio Tools menu, choose NuGet Package Manager, then Package Manager Console. To install the OWIN middleware libraries, enter the following commands in the Package Manager Console:

Powershell

```
Install-Package Microsoft.Owin.Security.OpenIdConnect
Install-Package Microsoft.Owin.Security.Cookies
Install-Package Microsoft.Owin.Host.SystemWeb
```

Next install the Microsoft Authentication Library with the following command:

Powershell

```
Install-Package Microsoft.Identity.Client -Pre
```


Finally install the Microsoft Graph Client Library with the following command:

Powershell

```
Install-Package Microsoft.Graph
```

## Back to coding

Now we're all set to do the sign in. Let's start by wiring the OWIN middleware to our app. Right-click the App_Start folder in Project Explorer and choose Add, then New Item. Choose the OWIN Startup Class template, name the file Startup.cs, and click Add. Replace the entire contents of this file with the following code.

C#

```
using System;
using System.Configuration;
using System.IdentityModel.Claims;
using System.IdentityModel.Tokens;
using System.Threading.Tasks;
using System.Web;

using Microsoft.IdentityModel.Protocols;
using Microsoft.Owin;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.Notifications;
using Microsoft.Owin.Security.OpenIdConnect;

using Owin;


[assembly: OwinStartup(typeof(dotnet_tutorial.App_Start.Startup))]

namespace dotnet_tutorial.App_Start
{
    public class Startup
    {
        public static string appId = ConfigurationManager.AppSettings["ida:AppId"];
        public static string appPassword = ConfigurationManager.AppSettings["ida:AppPassword"];
        public static string redirectUri = ConfigurationManager.AppSettings["ida:RedirectUri"];
        public static string[] scopes = ConfigurationManager.AppSettings["ida:AppScopes"]
          .Replace(' ', ',').Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

        public void Configuration(IAppBuilder app)
        {
            app.SetDefaultSignInAsAuthenticationType(CookieAuthenticationDefaults.AuthenticationType);

            app.UseCookieAuthentication(new CookieAuthenticationOptions());

            app.UseOpenIdConnectAuthentication(
              new OpenIdConnectAuthenticationOptions
              {
                  ClientId = appId,
                  Authority = "https://login.microsoftonline.com/common/v2.0",
                  Scope = "openid offline_access profile email " + string.Join(" ", scopes),
                  RedirectUri = redirectUri,
                  PostLogoutRedirectUri = "/",
                  TokenValidationParameters = new TokenValidationParameters
                  {
                      // For demo purposes only, see below
                      ValidateIssuer = false

                      // In a real multitenant app, you would add logic to determine whether the
                      // issuer was from an authorized tenant
                      //ValidateIssuer = true,
                      //IssuerValidator = (issuer, token, tvp) =>
                      //{
                      //  if (MyCustomTenantValidation(issuer))
                      //  {
                      //    return issuer;
                      //  }
                      //  else
                      //  {
                      //    throw new SecurityTokenInvalidIssuerException("Invalid issuer");
                      //  }
                      //}
                  },
                  Notifications = new OpenIdConnectAuthenticationNotifications
                  {
                      AuthenticationFailed = OnAuthenticationFailed,
                      AuthorizationCodeReceived = OnAuthorizationCodeReceived
                  }
              }
            );
        }

        private Task OnAuthenticationFailed(AuthenticationFailedNotification<OpenIdConnectMessage,
          OpenIdConnectAuthenticationOptions> notification)
        {
            notification.HandleResponse();
            string redirect = "/Home/Error?message=" + notification.Exception.Message;
            if (notification.ProtocolMessage != null && 
                !string.IsNullOrEmpty(notification.ProtocolMessage.ErrorDescription))
            {
                redirect += "&debug=" + notification.ProtocolMessage.ErrorDescription;
            }
            notification.Response.Redirect(redirect);
            return Task.FromResult(0);
        }

        private Task OnAuthorizationCodeReceived(AuthorizationCodeReceivedNotification notification)
        {
            notification.HandleResponse();
            notification.Response
                .Redirect("/Home/Error?message=See Auth Code Below&debug=" + notification.Code);
            return Task.FromResult(0);
        }
    }
}
```

Let's continue by adding a SignIn action to the HomeController class. Open the .\Controllers\HomeController.cs file. At the top of the file, add the following lines:

C#

```
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.OpenIdConnect;
using Microsoft.Graph;
```

Now add a new method called SignIn to the HomeController class.

SignIn action in ./Controllers/HomeController.cs

C#

```
public void SignIn()
{
    if (!Request.IsAuthenticated)
    {
        // Signal OWIN to send an authorization request to Azure
        HttpContext.GetOwinContext().Authentication.Challenge(
            new AuthenticationProperties { RedirectUri = "/" },
            OpenIdConnectAuthenticationDefaults.AuthenticationType);
    }
}
```


Finally, let's update the home page so that the login button invokes the SignIn action.

Updated contents of the ./Views/Home/Index.cshtml file

C#

```
@{
    ViewBag.Title = "Home Page";
}

<div class="jumbotron">
    <h1>ASP.NET MVC Tutorial</h1>
    <p class="lead">This sample app uses the Mail API to read messages in your inbox.</p>
    <p><a href="@Url.Action("SignIn", "Home", null, Request.Url.Scheme)" 
        class="btn btn-primary btn-lg">Click here to login</a></p>
</div>
```

Save your work and run the app. Click on the button to sign in. After signing in, you should be redirected to the error page, which displays an authorization code. Now let's do something with it.



