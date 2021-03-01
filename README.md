# msteams-dataverse-appt

## Prerequisites
1. Register App in Azure Active Directory and grant OnlineMeetings.ReadWrite.All API Application Permissions and generate a secret.  
2. Grant Users License to use Microsoft Teams.
3. Install [MicrosoftTeams Powershell Module] and connect.
4. [Grant Permissions] for the Azure AD Application to create [Online Meetings] on behalf of the users. 
5. Enable [Richt Text Editor for Appointments].

## Installation
1. Within \msteams-dataverse-appt\AddMsTeamsLinkToAppointmentPlugin\AddMsTeamsLinkToAppointmentPlugin\AddMsTeamsLinkToAppointmentActionHandler.cs populate the clientId, clientSecret and tenantId with the values obtained at Prerequisites point 1 
2. Compile Plugin.
3. Pack Solution and install it to your Dataverse environment. 

[MicrosoftTeams Powershell Module]: <https://docs.microsoft.com/en-us/microsoft-365/enterprise/manage-skype-for-business-online-with-microsoft-365-powershell?view=o365-worldwide>
[Grant Permissions]:<https://docs.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy>
[Online Meetings] : <https://docs.microsoft.com/en-us/graph/api/application-post-onlinemeetings?view=graph-rest-1.0&tabs=http>
[Richt Text Editor for Appointments]:<https://docs.microsoft.com/en-us/power-platform/admin/enable-rich-text-experience>
