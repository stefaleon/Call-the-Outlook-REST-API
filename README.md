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



## Exchanging the code for a token

Now let's update the OnAuthorizationCodeReceived function to retrieve a token. In ./App_Start/Startup.cs, add the following line to the top of the file:

C#

```
using Microsoft.Identity.Client;
```

Replace the current OnAuthorizationCodeReceived function with the following.

Updated OnAuthorizationCodeReceived action in ./App_Start/Startup.cs

C#

```
// Note the function signature is changed!
private async Task OnAuthorizationCodeReceived(AuthorizationCodeReceivedNotification notification)
{
    ConfidentialClientApplication cca = new ConfidentialClientApplication(
        appId, redirectUri, new ClientCredential(appPassword), null, null);

    string message;
    string debug;

    try
    {
        var result = await cca.AcquireTokenByAuthorizationCodeAsync(notification.Code, scopes);
        message = "See access token below";
        debug = result.AccessToken;
    }
    catch (MsalException ex)
    {
        message = "AcquireTokenByAuthorizationCodeAsync threw an exception";
        debug = ex.Message;
    }

    notification.HandleResponse();
    notification.Response.Redirect("/Home/Error?message=" + message + "&debug=" + debug);
}
```

Save your changes and restart the app. This time, after you sign in, you should see an access token displayed. 


Now that we can retrieve the access token, we need a way to save the token so the app can use it. The MSAL library defines a TokenCache interface, which we can use to store the tokens. Since this is a tutorial, we'll just implement a basic cache stored in the session.

Create a new folder in the project called TokenStorage. Add a class in this folder named SessionTokenCache, and replace the contents of that file with the following code (borrowed from https://github.com/Azure-Samples/active-directory-dotnet-webapp-openidconnect-v2).


Contents of the ./TokenCache/SessionTokenCache.cs file


C#

```
using System.Threading;
using System.Web;

using Microsoft.Identity.Client;

namespace dotnet_tutorial.TokenStorage
{
    // Adapted from https://github.com/Azure-Samples/active-directory-dotnet-webapp-openidconnect-v2
    public class SessionTokenCache
    {
        private static ReaderWriterLockSlim sessionLock = 
                                    new ReaderWriterLockSlim(LockRecursionPolicy.NoRecursion);
        string userId = string.Empty;
        string cacheId = string.Empty;
        HttpContextBase httpContext = null;

        TokenCache tokenCache = new TokenCache();

        public SessionTokenCache(string userId, HttpContextBase httpContext)
        {
            this.userId = userId;
            cacheId = userId + "_TokenCache";
            this.httpContext = httpContext;
            Load();
        }

        public TokenCache GetMsalCacheInstance()
        {
            tokenCache.SetBeforeAccess(BeforeAccessNotification);
            tokenCache.SetAfterAccess(AfterAccessNotification);
            Load();
            return tokenCache;
        }

        public bool HasData()
        {
            return (httpContext.Session[cacheId] != null && 
                                    ((byte[])httpContext.Session[cacheId]).Length > 0);
        }

        public void Clear()
        {
            httpContext.Session.Remove(cacheId);
        }

        private void Load()
        {
            sessionLock.EnterReadLock();
            tokenCache.Deserialize((byte[])httpContext.Session[cacheId]);
            sessionLock.ExitReadLock();
        }

        private void Persist()
        {
            sessionLock.EnterReadLock();

            // Optimistically set HasStateChanged to false. 
            // We need to do it early to avoid losing changes made by a concurrent thread.
            tokenCache.HasStateChanged = false;

            httpContext.Session[cacheId] = tokenCache.Serialize();
            sessionLock.ExitReadLock();
        }

        // Triggered right before ADAL needs to access the cache. 
        private void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            // Reload the cache from the persistent store in case it changed since the last access. 
            Load();
        }

        // Triggered right after ADAL accessed the cache.
        private void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            // if the access operation resulted in a cache update
            if (tokenCache.HasStateChanged)
            {
                Persist();
            }
        }
    }
}
```



Now let's update the OnAuthorizationCodeReceived function to use our new cache.

Add the following lines to the top of the Startup.cs file:


C#

```
using dotnet_tutorial.TokenStorage;
```


Replace the current OnAuthorizationCodeReceived function with this one.

New OnAuthorizationCodeReceived in ./App_Start/Startup.cs

C#

```
private async Task OnAuthorizationCodeReceived(AuthorizationCodeReceivedNotification notification)
{
    // Get the signed in user's id and create a token cache
    string signedInUserId = 
    notification.AuthenticationTicket.Identity.FindFirst(ClaimTypes.NameIdentifier).Value;
    SessionTokenCache tokenCache = new SessionTokenCache(signedInUserId, 
        notification.OwinContext.Environment["System.Web.HttpContextBase"] as HttpContextBase);

    ConfidentialClientApplication cca = new ConfidentialClientApplication(
        appId, redirectUri, new ClientCredential(appPassword), tokenCache.GetMsalCacheInstance(), null);

    try
    {
        var result = await cca.AcquireTokenByAuthorizationCodeAsync(notification.Code, scopes);
    }
    catch (MsalException ex)
    {
        string message = "AcquireTokenByAuthorizationCodeAsync threw an exception";
        string debug = ex.Message;
        notification.HandleResponse();
        notification.Response.Redirect("/Home/Error?message=" + message + "&debug=" + debug);
    }
}
```

Since we're saving stuff in the session, let's also add a SignOut action to the HomeController class. Add the following lines to the top of the HomeController.cs file:

C#

```
using dotnet_tutorial.TokenStorage;
```

Then add the SignOut action.

SignOut action in ./Controllers/HomeController.cs


C#

```
public void SignOut()
{
    if (Request.IsAuthenticated)
    {
        string userId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;

        if (!string.IsNullOrEmpty(userId))
        {
            // Get the user's token cache and clear it
            SessionTokenCache tokenCache = new SessionTokenCache(userId, HttpContext);
            tokenCache.Clear();
        }
    }
    // Send an OpenID Connect sign-out request. 
    HttpContext.GetOwinContext().Authentication.SignOut(
        CookieAuthenticationDefaults.AuthenticationType);
    Response.Redirect("/");
}
```

Now if you restart the app and sign in, you'll notice that the app redirects back to the home page, with no visible result. Let's update the home page so that it changes based on if the user is signed in or not.

First let's update the Index action in HomeController.cs to get the user's display name. Replace the existing Index function with the following code.


Updated Index action in ./Controllers/HomeController.cs

C#

```
public ActionResult Index()
{
    if (Request.IsAuthenticated)
    {
        string userName = ClaimsPrincipal.Current.FindFirst("name").Value;
        string userId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
        if (string.IsNullOrEmpty(userName) || string.IsNullOrEmpty(userId))
        {
            // Invalid principal, sign out
            return RedirectToAction("SignOut");
        }

        // Since we cache tokens in the session, if the server restarts
        // but the browser still has a cached cookie, we may be
        // authenticated but not have a valid token cache. Check for this
        // and force signout.
        SessionTokenCache tokenCache = new SessionTokenCache(userId, HttpContext);
        if (!tokenCache.HasData())
        {
            // Cache is empty, sign out
            return RedirectToAction("SignOut");
        }

        ViewBag.UserName = userName;
    }
    return View();
}
```

Next, replace the existing Index.cshtml code with the following.

Updated contents of ./Views/Home/Index.cshtml

C#

```
@{
    ViewBag.Title = "Home Page";
}

<div class="jumbotron">
    <h1>ASP.NET MVC Tutorial</h1>
    <p class="lead">This sample app uses the Mail API to read messages in your inbox.</p>
    @if (Request.IsAuthenticated)
    {
        <p>Welcome, @(ViewBag.UserName)!</p>
    }
    else
    {
        <p><a href="@Url.Action("SignIn", "Home", null, Request.Url.Scheme)" 
        class="btn btn-primary btn-lg">Click here to login</a></p>
    }
</div>
```

Finally, let's add some new menu items to the navigation bar if the user is signed in. Open the 
./Views/Shared/_Layout.cshtml file. Locate the following lines of code:

C#

```
<ul class="nav navbar-nav">
    <li>@Html.ActionLink("Home", "Index", "Home")</li>
    <li>@Html.ActionLink("About", "About", "Home")</li>
    <li>@Html.ActionLink("Contact", "Contact", "Home")</li>
</ul>
```

Replace those lines with this code:

C#

```
<ul class="nav navbar-nav">
    <li>@Html.ActionLink("Home", "Index", "Home")</li>
    @if (Request.IsAuthenticated)
    {
        <li>@Html.ActionLink("Inbox", "Inbox", "Home")</li>
        <li>@Html.ActionLink("Sign Out", "SignOut", "Home")</li>
    }
</ul>
```

Now if you save your changes and restart the app, the home page should display the logged on user's name and show two new menu items in the top navigation bar after you sign in. Now that we've got the signed in user and an access token, we're ready to call the Mail API.