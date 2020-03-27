// <copyright file="IGroupingHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common.Interfaces
{
    using System.Collections.Generic;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.GroupBot.Models;

    /// <summary>
    /// Handles Grouping Members with channel based on splitting criteria entered by group activity creator.
    /// </summary>
    public interface IGroupingHelper
    {
        /// <summary>
        /// Method sends grouping of members with channels in teams channel from where it is invoked.
        /// </summary>
        /// <param name="valuesFromTaskModule">Values enter by user in task module.</param>
        /// <param name="membersGroupingWithChannel">A dictionary of members grouped in channels.</param>
        /// <param name="groupActivityCreator">Team owner who initiated the group activity.</param>
        /// <returns>Returns a message to be shown in teams channels.</returns>
        string GroupingMessage(GroupDetail valuesFromTaskModule, Dictionary<int, IList<TeamsChannelAccount>> membersGroupingWithChannel, string groupActivityCreator);

        /// <summary>
        /// Method for grouping of members to channels based on specified number of groups.
        /// </summary>
        /// <param name="groupMembers">List of all members in channels except the team owners.</param>
        /// <param name="groupActivityCreator">Team owner who started the group activity.</param>
        /// <param name="channelCount">Number of channels to be created.</param>
        /// <param name="numberOfMembersInEachGroup">number of members in each group.</param>
        /// <returns>A dictionary with members grouped into channels based on entered grouping criteria.</returns>
        Dictionary<int, IList<TeamsChannelAccount>> SplitInGivenNumberOfGroups(IEnumerable<TeamsChannelAccount> groupMembers, TeamsChannelAccount groupActivityCreator, int channelCount, int numberOfMembersInEachGroup);

        /// <summary>
        /// Methods for grouping of members to channels based on given number of members for each group.
        /// </summary>
        /// <param name="groupMembers">Total channel members in the team.</param>
        /// <param name="groupActivityCreator">User who invoked the group create activity.</param>
        /// <param name="membersCount">Unit of number of members to be there in each group.</param>
        /// <returns>A dictionary with members grouped into channels based on entered grouping criteria.</returns>
        Dictionary<int, IList<TeamsChannelAccount>> SplitInGroupOfGivenMembers(IEnumerable<TeamsChannelAccount> groupMembers, TeamsChannelAccount groupActivityCreator, int membersCount);
    }
}