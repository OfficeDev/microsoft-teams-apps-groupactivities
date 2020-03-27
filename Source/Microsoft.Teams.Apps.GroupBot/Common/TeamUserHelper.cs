// <copyright file="TeamUserHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.GroupBot.Common.Interfaces;

    /// <summary>
    /// Class to get group members for grouping and verify if user is a team owner.
    /// </summary>
    public class TeamUserHelper : ITeamUserHelper
    {
        /// <summary>
        /// Helper for accessing with Microsoft Graph API.
        /// </summary>
        private readonly IGraphApiHelper graphApiHelper;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamUserHelper"/> class.
        /// </summary>
        /// <param name="graphApiHelper">Helper for accessing with Microsoft Graph API.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public TeamUserHelper(IGraphApiHelper graphApiHelper, ILogger<TeamUserHelper> logger)
        {
            this.graphApiHelper = graphApiHelper;
            this.logger = logger;
        }

        /// <summary>
        /// Method to check if user is a team owner.
        /// </summary>
        /// <param name="accessToken">Microsoft Graph API access token.</param>
        /// <param name="teamGroupId">Azure Active Directory (AAD) Group id.</param>
        /// <param name="userId">Users object ID within Azure Active Directory (AAD). </param>
        /// <returns>Returns true if user is team owner and returns null if exception.</returns>
        public async Task<bool?> VerifyIfUserIsTeamOwnerAsync(string accessToken, string teamGroupId, string userId)
        {
            try
            {
                var teamOwners = await this.graphApiHelper.GetOwnersAsync(accessToken, teamGroupId);
                return teamOwners?.TeamOwnerValues.Any(teamOwner => teamOwner.TeamOwnerId.ToString().Contains(userId));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while checking if the user is team owner for user AadObject : {userId}");
                return null;
            }
        }

        /// <summary>
        /// Method to get list of members for grouping except the owners of team.
        /// </summary>
        /// <param name="accessToken ">Token to access Microsoft Graph API.</param>
        /// <param name="teamGroupId">groupId of the team in which channel is to be created.</param>
        /// <param name="teamMembers">List of all team members in a channel.</param>
        /// <returns>A task that returns list of members for grouping.</returns>
        public async Task<IEnumerable<TeamsChannelAccount>> GetGroupMembersAsync(string accessToken, string teamGroupId, IEnumerable<TeamsChannelAccount> teamMembers)
        {
            // Logs total no of members in a team.
            this.logger.LogInformation($"Total number of users in team is : {teamMembers.Count()}");

            try
            {
                var teamOwners = await this.graphApiHelper.GetOwnersAsync(accessToken, teamGroupId);
                if (teamOwners == null)
                {
                    this.logger.LogInformation($"Team Owners obtained is null for teamGroupID : {teamGroupId}");
                    return null;
                }

                // List of all members in channels except the owners of team.
                var groupMembers = from member in teamMembers
                                   where !(from owner in teamOwners?.TeamOwnerValues select owner.TeamOwnerId.ToString()).Contains(member.AadObjectId)
                                   select member;

                // Logs total no of members in a group activity.
                this.logger.LogInformation($"Total number of users in group activity is : {groupMembers.Count()}");

                return groupMembers;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while getting group members for performing grouping for teamGroupId: {teamGroupId}. ");
                return null;
            }
        }
    }
}
