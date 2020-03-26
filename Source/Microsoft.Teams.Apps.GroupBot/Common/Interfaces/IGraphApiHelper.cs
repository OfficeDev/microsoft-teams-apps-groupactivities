// <copyright file="IGraphApiHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common.Interfaces
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GroupBot.Models;
    using Microsoft.Teams.Apps.GroupBot.Models.ChannelListDetails;
    using Microsoft.Teams.Apps.GroupBot.Models.TeamOwnerDetails;

    /// <summary>
    /// Provides the helper methods to access Microsoft Graph API.
    /// </summary>
    public interface IGraphApiHelper
    {
        /// <summary>
        /// Get team owners list from Microsoft Graph.
        /// </summary>
        /// <param name="token">Microsoft Graph user access token.</param>
        /// <param name="groupId">Team Azure Active Directory object id of the channel where bot is installed.</param>
        /// <returns>A task that returns team owners.</returns>
        Task<TeamOwnerDetails> GetOwnersAsync(string token, string groupId);

        /// <summary>
        /// Create public channel using Microsoft Graph API.
        /// </summary>
        /// <param name="token">Microsoft Graph user access token.</param>
        /// <param name="body">Group activity details entered by user in task module.</param>
        /// <param name="groupId">Azure Active Directory (AAD) Group Id for the team.</param>
        /// <returns>A task returns true if channel is successfully created else false.</returns>
        Task<ChannelApiResponse> CreatePublicChannelAsync(string token, string body, string groupId);

        /// <summary>
        /// Create private channel using Microsoft Graph API.
        /// </summary>
        /// <param name="token">Microsoft Graph user access token.</param>
        /// <param name="body">Group activity details entered by user in task module.</param>
        /// <param name="groupId">Azure Active Directory (AAD) Group Id for the team.</param>
        /// <returns>A task returns true if channel is successfully created else false.</returns>
        Task<ChannelApiResponse> CreatePrivateChannelAsync(string token, string body, string groupId);

        /// <summary>
        /// Get list of all channels in a team.
        /// </summary>
        /// <param name="token">Azure Active Directory (AAD) token to access graph API.</param>
        /// <param name="groupId">groupId of the team in which channel is to be created.</param>
        /// <returns>A task that returns list of all channels in a team.</returns>
        Task<ChannelListRequest> GetChannelsAsync(string token, string groupId);
    }
}