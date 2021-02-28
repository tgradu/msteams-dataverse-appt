using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace AddMsTeamsLinkToAppointmentPlugin
{
    [DataContract]
    public class OnlineMeetingInformation
    {
        [DataMember(Name = "startDateTime")]
        public string StartDateTime { get; set; } = DateTimeOffset.UtcNow.ToString("O");

        [DataMember(Name = "endDateTime")]
        public string EndDateTime { get; set; } = DateTimeOffset.UtcNow.AddMinutes(30).ToString("O");

        [DataMember(Name = "subject")]
        public string Subject { get; set; } = "Online Meeting";

        [DataMember(Name = "@odata.type")]
        public string OdataType { get; set; } = "microsoft.graph.onlineMeeting";

        [DataMember(Name = "joinInformation")]
        public JoinInformationData JoinInformationData { get; set; }
    }

    [DataContract]
    public class JoinInformationData
    {
        [DataMember(Name = "content")]
        public string Content { get; set; }
    }

    public class OnlineMeetingGenerator
    {
        public OnlineMeetingInformation GenerateOnlineMeeting(OnlineMeetingInformation onlineMeetingRequest, AccessToken token, Guid azureAdUserObjectId)
        {
            Uri graphEndpoint = new Uri($"https://graph.microsoft.com/v1.0/users/{azureAdUserObjectId}/onlineMeetings");
            string meetingJson = SerializeOnlineMeeting(onlineMeetingRequest);

            HttpClient httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue(token.Type, token.Value);

            HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, graphEndpoint);
            httpRequestMessage.Content = new StringContent(meetingJson, Encoding.UTF8, "application/json");

            var response = httpClient.SendAsync(httpRequestMessage).GetAwaiter().GetResult();

            if (!response.IsSuccessStatusCode)
                throw new Exception($"Create Online Meeting failed : {response.ReasonPhrase}");

            DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings()
            {
                UseSimpleDictionaryFormat = true
            };

            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(OnlineMeetingInformation), settings);
            var onlineMeetindDeserialized = (OnlineMeetingInformation)ser.ReadObject(response.Content.ReadAsStreamAsync().GetAwaiter().GetResult());

            return onlineMeetindDeserialized;

        }

        private string SerializeOnlineMeeting(OnlineMeetingInformation onlineMeeting)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings()
                {
                    UseSimpleDictionaryFormat = true
                };
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(OnlineMeetingInformation), settings);

                ser.WriteObject(ms, onlineMeeting);

                using (StreamReader sr = new StreamReader(ms))
                {
                    ms.Position = 0;
                    return sr.ReadToEnd();
                }
            }
        }
    }
}
