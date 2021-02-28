using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace AddMsTeamsLinkToAppointmentPlugin
{
    [DataContract]
    public class AccessToken
    {
        [DataMember(Name = "token_type")]
        public string Type { get; set; }

        [DataMember(Name = "access_token")]
        public string Value { get; set; }
    }

    public class TokenProvider
    {
        public AccessToken GetToken(string clientId, string clientSecret, string tenantId, string scope)
        {
            var httpClient = new HttpClient();

            HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, new Uri($"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token"));
            httpRequestMessage.Content = new FormUrlEncodedContent(new KeyValuePair<string, string>[]
            {
                new KeyValuePair<string,string>( "client_id", clientId),
                new KeyValuePair<string,string>( "client_secret", clientSecret),
                new KeyValuePair<string,string>( "scope", scope),
                new KeyValuePair<string,string>( "grant_type", "client_credentials"),
            });

            var result = httpClient.SendAsync(httpRequestMessage).GetAwaiter().GetResult();

            DataContractJsonSerializerSettings settings = new DataContractJsonSerializerSettings()
            {
                UseSimpleDictionaryFormat = true
            };

            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(AccessToken), settings);
            var tokenResponse = (AccessToken)ser.ReadObject(result.Content.ReadAsStreamAsync().GetAwaiter().GetResult());

            return tokenResponse;
        }
    }
}
