# recoverlette
quickly rewrite job application cover letter Word documents and export as PDF

# Why and Wherefore
Microsoft (and every other software vendor, of course, but our tool requires attention to how Microsoft is doing things, so they get called out) has been slowly and steadily forcing users into the Cloud (specifically, their Cloud - that is, Azure) by making it increasingly difficult with every minor Office update to manage your own files on your local disk, making it so uncomfortable to do what is necessary to avoid using OneDrive, to try to keep control and ownership of your own data, to set your own privacy policy, or generally just edit and save text in a ubiquitous file format using your own naming schemes and opting for whatever revision control or file management scheme you prefer, that fatigue from it all will (they hope) wear you down to a point of surrender. A victory for them where the defeated don't have the capability to buy a version of a word processor to run locally on your machine, editing files locally on your machine, and _stick with that version if that is your wish_. Nope, no more one-time purchases, but a dystopic everything-as-a-service, everything in the cloud, online security nightmare, with everything offered only in a recurring, auto-renewing paid subscription model. Just what we always wanted. Rejoice.

Well, friend, let me tell you... I was right there with you. I was ready to give in and surrender in defeat.  The stakes were high. I was actively applying for jobs and even after giving in and reactivating a microsoft logon so I could use the Office365 family subscription my wife signed us up for while she was in law school, and then capitulating and enabling ONeDrive so that when I use the file upload dialog in the Office portal, the files I upload actually go somewhere that I can access them (yeah, prior to that the files seemingly uploaded successfully insofar as the http file upload dialog indicated, but the files themselves were nowhere to be found when using Word. *shrug*), after resigning to having to accept their workflow and revision control, and auto-saving, and all such nonsense (there may be some way to be a Word "power user" and learn configfuration tips and tricks to make this experience less shitty, but Office productivity app expertise is just not something I have much desire to invest a whole lot of time on searching reading complaining etc. It's a me problem I know.)

So even after all of that acquiesence to the MS Borg, I found the process of actually using the O365 Word interface to be embiggenly enshittified over the old local Windows application. And due to the slow and steady feature migration and such we also have an inconsistent experience from Windows to Mac, or efrom one version to the next if installed locally, even on two different Macs. I gave in and accepted their cloudfuckery and find there are features that haven't made it there yet. Again let us rejoice in the blessings from our bountiful gifts. Additionallythe implementation of Styles, for example, is a - for lack of a better word - _destructive_ operation from which there is no undo. Oh and thanks to the timely autosaving feature, it's permanent for all practical purposes. Paranoia had me implementing an ad hoc brutish version control by exporting the file (that is, downloading a local copy) with some name with a string in the filename identifying its revision to me at a glance, so that I could revert to a known good state after experiencing the non-intuitive mayhem that Styles box can unleash. (Again, I am self-aware enough to know that it is more than likely that I am a shitty business productivity app user, and that other people love these "enhancements.")

But I do not love them, not at all. I long ago ceded to the fact that I can't just keep an ascii text file - maybe even marked up with html or markdown - edit it in vim, or better yet, from the command-line, sedding and awking it just right before /usr/ucb/mailing it do the hiring manager. I know this is silly to even wish for, as the presentation - the fonts and such - tidiness afforded us by the likes of Microsoft and Adobe file formatting can look just great and (most) things can usually (or used to be) done in a mostly intuitive manner without having to read manuals or stufy the internals.

[to be cont'd]


# Installation
```
pip install -r requirements.txt
```

# Usage
## Prerequisites
### For personal Office 365 authentication and API access:

Instead of Azure AD, you'll use the Microsoft Identity Platform (formerly known as Microsoft Account) for authentication.

- Register your application in the Microsoft Application Registration Portal 
    - Go to https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
    - Click "New registration"
    - Name your app and select "Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"
    - For Redirect URI, select "Public client/native (mobile & desktop)" and enter "http://localhost"
    - After registration, note down the Application (client) ID

- Set up authentication using the OAuth 2.0 authorization code flow.


- Use the Microsoft Graph SDK for Python to interact with your files.

# TODO
- Add ability to replace entire text body
# TODO (possibly)
- Support modifying font, font size, font color?
- Add support for AAD and AAD application for those who want to send resumes using their corporate Enterprise user for some reason
    (using office365 REST API? SharePoint/OneDrive? Haven't started looking into this yet; just jotting thoughts, but it's what I looked at before settling  on msgraph API)
- Add support for certs instead of user credentials
- 

