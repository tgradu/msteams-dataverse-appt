using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Net;

namespace AddMsTeamsLinkToAppointmentPlugin
{
    public class AddMsTeamsLinkToAppointmentActionHandler : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            string clientId = "";
            string clientSecret = "";
            string tenantId = "";
            string scope = "https://graph.microsoft.com/.default";

            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var orgService = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                var tokenProvider = new TokenProvider();
                var token = tokenProvider.GetToken(clientId, clientSecret, tenantId, scope);

                Entity systemUser = orgService.Retrieve("systemuser", context.UserId, new ColumnSet("azureactivedirectoryobjectid"));
                Guid userAzureAdObjectId = systemUser.GetAttributeValue<Guid>("azureactivedirectoryobjectid");
                OnlineMeetingGenerator onlineMeetingGenerator = new OnlineMeetingGenerator();
                OnlineMeetingInformation onlineMeetingInformation = new OnlineMeetingInformation();

                var onlineMeetingResponse = onlineMeetingGenerator.GenerateOnlineMeeting(onlineMeetingInformation, token, userAzureAdObjectId);

                var meetingContent = WebUtility.UrlDecode(onlineMeetingResponse.JoinInformationData.Content).Remove(0, 15);

                context.OutputParameters.AddOrUpdateIfNotNull("OnlineMeeting", meetingContent);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
