# OutlookSignature

Usecase:
Bulk set a standard template Signature in a large organization by deploying this tiny app.

Simple C# Win Forms application which:
-Gets an Outlook Signature Template from a shared folder (1 sub folder and 3 files)
-Gets user data from Active Directory
-Creates a new Outlook Signature in the user's directory by replacing variables in the Template with actual user data
