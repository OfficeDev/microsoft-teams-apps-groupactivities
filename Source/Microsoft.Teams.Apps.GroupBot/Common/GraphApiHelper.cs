// <copyright file="GraphApiHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common
{
    using System;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.GroupBot.Common.Interfaces;
    using Microsoft.Teams.Apps.GroupBot.Models;
    using Microsoft.Teams.Apps.GroupBot.Models.ChannelListDetails;
    using Microsoft.Teams.Apps.GroupBot.Models.TeamOwnerDetails;
    using Newtonsoft.Json;

    /// <summary>
    /// The class that represent the helper methods to access Microsoft Graph API.
    /// </summary>
    public class GraphApiHelper : IGraphApiHelper
    {
        /// <summary>
        /// Provides a base class for sending HTTP requests and receiving HTTP responses from a resource identified by a URI.
        /// </summary>
        /// </summary>
        private readonly HttpClient client;

        /// <summary>
        /// Instance to send logs to the Application Insights service..
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphApiHelper"/> class.
        /// </summary>
        /// <param name="client">Provides a base class for sending HTTP requests and receiving HTTP responses from a resource identified by a URI.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public GraphApiHelper(HttpClient client, ILogger<GraphApiHelper> logger)
        {
            this.client = client;
            this.logger = logger;
        }

        /// <summary>
        /// Get owner list from Microsoft Graph.
        /// </summary>
        /// <param name="token">Microsoft Graph user access token.</param>
        /// <param name="groupId">groupId of the team in which channel is to be created.</param>
        /// <returns>A task that returns team owner details.</returns>
        public async Task<TeamOwnerDetails> GetOwnersAsync(string token, string groupId)
        {
            var response = await this.GetAsync(token, $"{Constants.GraphAPIBaseURL}/v1.0/groups/{groupId}/owners");

            if (response.IsSuccessStatusCode)
            {
                this.logger.LogInformation($"Graph API call to get owners successful with statusCode - {response.StatusCode}");
                return await this.DeserializeJsonStringAsync<TeamOwnerDetails>(response);
            }

            var errorMessage = await response.Content.ReadAsStringAsync();
            this.logger.LogInformation($"Graph API call to get owners error - {errorMessage} statusCode - {response.StatusCode}");
            return null;
        }

        /// <summary>
        /// Create public channel using Microsoft Graph API.
        /// </summary>
        /// <param name="token">Azure Active Directory (AAD) token to access graph API.</param>
        /// <param name="body">Body to be sent to API.</param>
        /// <param name="groupId">groupId of the team in which channel is to be created.</param>
        /// <returns>A task that returns response with created public channel details.</returns>
        public async Task<ChannelApiResponse> CreatePublicChannelAsync(string token, string body, string groupId)
        {
            try
            {
                string requestUrl = $"{Constants.GraphAPIBaseURL}/v1.0/teams/{groupId}/channels";
                var response = await this.PostAsync(token, body, requestUrl);
                if (response.IsSuccessStatusCode)
                {
                    this.logger.LogInformation($"Graph API call to create public channel is successful with statusCode - {response.StatusCode}");
                    return await this.DeserializeJsonStringAsync<ChannelApiResponse>(response);
                }

                var errorMessage = await response.Content.ReadAsStringAsync();
                this.logger.LogWarning($"Graph API call to create public channel error- {errorMessage} statusCode - {response.StatusCode}");
                return null;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Graph API create public channel error");
                return null;
            }
        }

        /// <summary>
        /// Create private channel using Microsoft Graph API.
        /// </summary>
        /// <param name="token">Azure Active Directory (AAD) token to access graph API.</param>
        /// <param name="body">Body to be sent to API.</param>
        /// <param name="groupId">groupId of the team in which channel is to be created.</param>
        /// <returns>A task that returns response with created private channel details.</returns>
        public async Task<ChannelApiResponse> CreatePrivateChannelAsync(string token, string body, string groupId)
        {
            try
            {
                string requestUrl = $"{Constants.GraphAPIBaseURL}/beta/teams/{groupId}/channels";
                var response = await this.PostAsync(token, body, requestUrl);

                if (response.IsSuccessStatusCode)
                {
                    this.logger.LogInformation($"Graph API call to create private channel is successful with statusCode - {response.StatusCode}");
                    return await this.DeserializeJsonStringAsync<ChannelApiResponse>(response);
                }

                var errorMessage = await response.Content.ReadAsStringAsync();
                this.logger.LogWarning($"Graph API create private channel error- {errorMessage} statusCode - {response.StatusCode}");
                return null;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Graph API create private channel error");
                return null;
            }
        }

        /// <summary>
        /// Get list of all channels in a team.
        /// </summary>
        /// <param name="token">Azure Active Directory (AAD) token to access graph API.</param>
        /// <param name="groupId">groupId of the team in which channel is to be created.</param>
        /// <returns>A task that returns list of all channels in a team.</returns>
        public async Task<ChannelListRequest> GetChannelsAsync(string token, string groupId)
        {
            var response = await this.GetAsync(token, $"{Constants.GraphAPIBaseURL}/beta/teams/{groupId}/channels");

            if (response.IsSuccessStatusCode)
            {
                this.logger.LogInformation($"Graph API call to get list of channels is successful with statusCode - {response.StatusCode}");
                return await this.DeserializeJsonStringAsync<ChannelListRequest>(response);
            }

            var errorMessage = await response.Content.ReadAsStringAsync();
            this.logger.LogInformation($"Graph API get channels error- {errorMessage} statusCode - {response.StatusCode}");
            return null;
        }

        /// <summary>
        /// Method to get data from API.
        /// </summary>
        /// <param name="token">Microsoft Graph API user access token.</param>
        /// <param name="requestUrl">Microsoft Graph API request URL.</param>
        /// <returns>A task that represents a HTTP response message including the status code and data.</returns>
        private async Task<HttpResponseMessage> GetAsync(string token, string requestUrl)
        {
            HttpMethod httpMethod = new HttpMethod("GET");
            var request = new HttpRequestMessage(httpMethod, requestUrl);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

            return await this.client.SendAsync(request);
        }

        /// <summary>
        /// Method to post data to API.
        /// </summary>
        /// <param name="token">Microsoft Graph API user access token.</param>
        /// <param name="body">Body to be sent to API.</param>
        /// <param name="requestUrl">Microsoft Graph API request URL.</param>
        /// <returns>A task that represents a HTTP response message including the status code and data.</returns>
        private async Task<HttpResponseMessage> PostAsync(string token, string body, string requestUrl)
        {
            HttpMethod httpMethod = new HttpMethod("POST");
            var request = new HttpRequestMessage(httpMethod, requestUrl);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            this.client.DefaultRequestHeaders.Remove("Authorization");
            this.client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
            request.Content = new StringContent(body, Encoding.UTF8, "application/json");

            return await this.client.SendAsync(request);
        }

        /// <summary>
        /// De-serialize the response from HTTP response data.
        /// </summary>
        /// <typeparam name="T">Model to De-serialize data.</typeparam>
        /// <param name="response">Represents a HTTP response message including the status code and data.</param>
        /// <returns>De-serialized HTTP response data.</returns>
        private async Task<T> DeserializeJsonStringAsync<T>(HttpResponseMessage response)
            where T : class
        {
            return JsonConvert.DeserializeObject<T>(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
        }
    }
}
