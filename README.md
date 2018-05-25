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





