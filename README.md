# Get Excel workbooks using Microsoft Graph and MSAL in an Office Add-in 

Learn how to build a Microsoft Office Add-in that connects to Microsoft Graph, finds the first three workbooks stored in OneDrive for Business, fetches their filenames, and inserts the names into a range on an Excel worksheet using Office.js.

# Introduction

Integrating data from online service providers increases the value and adoption of your add-ins. This code sample shows you how to connect your add-in to Microsoft Graph. Use this code sample to:

* Connect to Microsoft Graph from an Office Add-in.
* Use the MSAL .NET Library to implement the OAuth 2.0 authorization framework in an add-in.
* Use the OneDrive REST APIs from Microsoft Graph.
* Show a dialog using the Office UI namespace.
* Build an Add-in using ASP.NET MVC, MSAL 3.x.x for .NET,  and Office.js. 
* Use add-in commands in an add-in.


## Prerequisites
To run this code sample, the following are required.

* Visual Studio 2019 or later.

* SQL Server Express (No longer automatically installed with recent versions of Visual Studio.)

* An Office 365 account which you can get by joining the [Office 365 Developer Program](https://aka.ms/devprogramsignup) that includes a free 1 year subscription to Office 365.

* At least three Excel workbooks stored on OneDrive for Business in your Office 365 subscription.

* Excel for Windows, version 16.0.6769.2001 or higher.
* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* A Microsoft Azure Tenant. This add-in requires Azure Active Directiory (AD). Azure AD provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Configure the project

1. In **Visual Studio**, choose the **Office-Add-in-Microsoft-Graph-ASPNETWeb** project. In **Properties**, ensure **SSL Enabled** is **True**. Verify that the **SSL URL** uses the same domain name and port number as those listed in the next step.
 
2. Register your application using the [Azure Management Portal](https://manage.windowsazure.com). **Log in with the identity of an administrator of your Office 365 tenancy to ensure that you are working in an Azure Active Directory that is associated with that tenancy.** To learn how to register your application, see [Register an application with the Microsoft Identity Platform](https://docs.microsoft.com/graph/auth-register-app-v2). Use the following settings:

 - REDIRCT URI: https://localhost:44301/AzureADAuth/Authorize	
 - SUPPORTED ACCOUNT TYPES: "Accounts in this organizational directory only"
 - IMPLICIT GRANT: Do not enable any Implicit Grant options
 - API PERMISSIONS: **Files.Read.All** and **User.Read**

	> Note: After you register your application, copy the **Application (client) ID** and the **Directory (tenant) ID** on the **Overview** blade of the App Registration in the Azure Management Portal. When you create the client secret on the **Certificates & secrets** blade, copy it too. 
	 
3.  In web.config, use the values that you copied in the previous step. Set **AAD:ClientID** to your client id, set **AAD:ClientSecret** to your client secret, and set **"AAD:O365TenantID"** to your tenant ID. 

## Run the project
1. Open the Visual Studio solution file. 
2. Right-click **Office-Add-in-Microsoft-Graph-ASPNET**, and then choose **Set as StartUp Project**.
2. Press F5. 
3. In PowerPoint, choose **Insert** > **Open Files** in the **OneDrive Files** group to open the task pane add-in.

## Known issues

* The Fabric spinner control appears only briefly or not at all. 

* After you logout in the task pane, you will briefly see, in the task pane, the standard AAD page that confirms that you have logged out and recommends that you close all browser windows. Usually, this lasts only a brief moment before you are redirected to the add-ins home page. Sometimes the redirection never happens and you must reload the add-in from the personality menu.

* Scenario: When trying to run the code sample, the add-in will not load.
	* Resolution: 
		1. In Visual Studio, open **SQL Server Object Explorer**.
		2. Expand **(localdb)\MSSQLLocalDB** > **Databases**.
		3. Right click **Office-Add-in-Microsoft-Graph-ASPNET**, then choose **Delete**. 
* Scenario: When you run the code sample, you get an error on the line *Office.context.ui.messageParent*.	
	* Resolution: Stop running the code sample and restart it. 
* If download the zip file, when you extract the files you get an error indicating that the file path is too long.
	* Resolution: Unzip your files to a folder directly under the root (e.g. c:\sample).

## Questions and comments
We'd love to get your feedback about the *Get Excel workbooks using Microsoft Graph and MSAL in an Office Add-in* sample. You can send your feedback to us in the *Issues* section of this repository.
Questions about Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Ensure your questions are tagged with [office-js], [MicrosoftGraph] and [API].

## Additional resources

* [Microsoft Graph (Excel) ToDo code sample](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-ExcelREST-ToDo)
* [Microsoft Graph documentation](https://docs.microsoft.com/graph/)
* [Office Add-ins documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-inss)

## Copyright
Copyright (c) 2019 Microsoft Corporation. All rights reserved.



This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
