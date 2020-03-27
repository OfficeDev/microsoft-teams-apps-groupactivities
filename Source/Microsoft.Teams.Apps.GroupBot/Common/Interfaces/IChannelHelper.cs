// <copyright file="IChannelHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common.Interfaces
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.GroupBot.Models;

    /// <summary>
    /// Handles creating channels based on grouping.
    /// </summary>
    public interface IChannelHelper
    {
        /// <summary>
        /// Validate channel count and create channel using Microsoft Graph API.
        /// </summary>
        /// <param name="token">Token to access Microsoft Graph API.</param>
        /// <param name="groupActivityId">Guid group activity Id.</param>
        /// <param name="teamId">Team id where messaging extension is installed.</param>
        /// <param name="groupId">Team Azure Active Directory object id of the channel where bot is installed.</param>
        /// <param name="groupingMessage">Grouping message with members mapped to groups.</param>
        /// <param name="valuesFromTaskModule">Group activity details obtained from task module as entered by user.</param>
        /// <param name="membersGroupingWithChannel">List of all members grouped in channels based on grouping criteria.</param>
        /// <param name="groupActivityCreator">Team owner who started the group activity.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that returns true if channel is created successfully.</returns>
        Task ValidateAndCreateChannelAsync(string token, string groupActivityId, string teamId, string groupId, string groupingMessage, GroupDetail valuesFromTaskModule, Dictionary<int, IList<TeamsChannelAccount>> membersGroupingWithChannel, TeamsChannelAccount groupActivityCreator, ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken);

        /// <summary>
        /// Method stores group activity detail to azure table storage.
        /// </summary>
        /// <param name="serviceUrl">Bot activity service URL.</param>
        /// <param name="groupActivityId">Group activity Id.</param>
        /// <param name="teamId">Team id where messaging extension is installed.</param>
        /// <param name="valuesFromTaskModule">Group activity details obtained from task module as entered by user.</param>
        /// <param name="groupActivityCreator">Team owner who started the group activity.</param>
        /// <param name="groupingCardConversationId">Conversation id of grouping card posted in channel.</param>
        /// <param name="groupingCardActivityId">Activity id of grouping card posted in channel.</param>
        /// <returns>Returns a task that stores group activity to azure table storage. </returns>
        Task StoreGroupActivityDetailsAsync(string serviceUrl, string groupActivityId, string teamId, GroupDetail valuesFromTaskModule, string groupActivityCreator, string groupingCardConversationId, string groupingCardActivityId);
    }
}