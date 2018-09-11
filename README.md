# VSTS Attachement Downloader

Using `Microsoft.TeamFoundationServer.Client`, `Microsoft.VisualStudio.Services.Client` and `Microsoft.VisualStudio.Services.InteractiveClient`, this console app download work item's attachement by accessing VSTS REST APIs

You need to register an Azure app in your Azure portal to get a client id, and set the reply url (could be any url e.g. http://vstsdownloader), then put the same information into Program.cs file.
