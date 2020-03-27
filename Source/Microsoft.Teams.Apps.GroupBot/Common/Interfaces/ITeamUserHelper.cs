// <copyright file="ITeamUserHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Bot.Schema.Teams;

    /// <summary>
    /// Handles fetching group members for grouping and verify if user is a team owner.
    /// </summary>
    public interface ITeamUserHelper
    {
        /// <summary>
        /// Method to get list of members for grouping except the owners of team.
        /// </summary>
        /// <param name="accessToken">Token to access Microsoft Graph API.</param>
        /// <param name="teamGroupId">groupId of the team in which channel is to be created.</param>
        /// <param name="teamMembers">List of all team members in a channel.</param>
        /// <returns>A task that returns list of members for grouping.</returns>
        Task<IEnumerable<TeamsChannelAccount>> GetGroupMembersAsync(string accessToken, string teamGroupId, IEnumerable<TeamsChannelAccount> teamMembers);

        /// <summary>
        /// Method to check if user is a team owner.
        /// </summary>
        /// <param name="accessToken">Microsoft Graph API access token.</param>
        /// <param name="teamGroupId">Azure Active Directory (AAD) Group id.</param>
        /// <param name="userId">User object ID within Azure Active Directory (AAD). </param>
        /// <returns>Returns true if user is team owner and returns null if exception.</returns>
        Task<bool?> VerifyIfUserIsTeamOwnerAsync(string accessToken, string teamGroupId, string userId);
    }
}