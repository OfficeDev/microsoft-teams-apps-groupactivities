// <copyright file="GroupingHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.GroupBot.Common.Interfaces;
    using Microsoft.Teams.Apps.GroupBot.Models;

    /// <summary>
    /// Group Members with channel based on splitting criteria entered by group activity creator.
    /// </summary>
    public class GroupingHelper : IGroupingHelper
    {
        /// <summary>
        /// Maximum length of group name to be shown in grouping message.
        /// </summary>
        private const int TruncateThresholdLength = 40;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="GroupingHelper"/> class.
        /// Class that handles grouping of members with channel.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public GroupingHelper(ILogger<GroupingHelper> logger)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Methods for grouping of members to channels based on given number of members for each group.
        /// </summary>
        /// <param name="groupMembers">Total channel members in the team.</param>
        /// <param name="groupActivityCreator">User who invoked the group create activity.</param>
        /// <param name="membersCount">Unit of number of members to be there in each group.</param>
        /// <returns>A dictionary with members grouped into channels based on entered grouping criteria.</returns>
        public Dictionary<int, IList<TeamsChannelAccount>> SplitInGroupOfGivenMembers(IEnumerable<TeamsChannelAccount> groupMembers, TeamsChannelAccount groupActivityCreator, int membersCount)
        {
            try
            {
                Dictionary<int, IList<TeamsChannelAccount>> membersGroupingWithChannel = new Dictionary<int, IList<TeamsChannelAccount>>();
                var teamsChannels = new List<TeamsChannelAccount>();
                int numberOfMembersInEachGroupCount = 0, channelIndex = 0;
                var random = new Random();
                var randomlyOrderedGroupMembers = groupMembers.OrderBy(i => random.Next());

                foreach (var member in randomlyOrderedGroupMembers)
                {
                    // Add owner to channel once required number of members are added to channel.
                    if (numberOfMembersInEachGroupCount == membersCount)
                    {
                        teamsChannels.Add(groupActivityCreator);
                        membersGroupingWithChannel[channelIndex] = teamsChannels.ToList();
                        channelIndex++;
                        numberOfMembersInEachGroupCount = 0;
                        teamsChannels.Clear();
                    }

                    // Add members to dictionary.
                    teamsChannels.Add(member);
                    numberOfMembersInEachGroupCount++;
                }

                // Add remaining members and group creator to dictionary .
                if (teamsChannels.Count > 0)
                {
                    teamsChannels.Add(groupActivityCreator);
                    membersGroupingWithChannel[channelIndex] = teamsChannels.ToList();
                }

                return membersGroupingWithChannel;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while performing grouping logic based on given number of members in each group.");
                return null;
            }
        }

        /// <summary>
        /// Method for grouping of members to channels based on specified number of groups.
        /// </summary>
        /// <param name="groupMembers">List of all members in channels except the team owners.</param>
        /// <param name="groupActivityCreator">Team owner who started the group activity.</param>
        /// <param name="channelCount">Number of channels to be created.</param>
        /// <param name="numberOfMembersInEachGroup">number of members in each group.</param>
        /// <returns>A dictionary with members grouped into channels based on entered grouping criteria.</returns>
        public Dictionary<int, IList<TeamsChannelAccount>> SplitInGivenNumberOfGroups(IEnumerable<TeamsChannelAccount> groupMembers, TeamsChannelAccount groupActivityCreator, int channelCount, int numberOfMembersInEachGroup)
        {
            try
            {
                var teamsChannels = new List<TeamsChannelAccount>();
                Dictionary<int, IList<TeamsChannelAccount>> channelsGroupingwithMembers = new Dictionary<int, IList<TeamsChannelAccount>>();
                int numberOfMembersCount = 0, numberofGroupsIndex = 0;

                var random = new Random();
                var randomlyOrderedGroupMembers = groupMembers.OrderBy(i => random.Next());

                // Add each member at respective index in dictionary, where index represents a channel.
                foreach (var member in randomlyOrderedGroupMembers)
                {
                    if (numberOfMembersCount == numberOfMembersInEachGroup && numberofGroupsIndex != channelCount)
                    {
                        teamsChannels.Add(groupActivityCreator);
                        channelsGroupingwithMembers[numberofGroupsIndex] = teamsChannels.ToList();
                        numberOfMembersCount = 0;
                        numberofGroupsIndex++;
                        teamsChannels.Clear();
                    }

                    teamsChannels.Add(member);
                    numberOfMembersCount++;
                }

                // Add the remaining team members to dictionary including team owner.
                if (teamsChannels.Count > 0)
                {
                    if (randomlyOrderedGroupMembers.Count() % channelCount == 0)
                    {
                        teamsChannels.Add(groupActivityCreator);
                        channelsGroupingwithMembers[numberofGroupsIndex] = teamsChannels.ToList();
                    }
                    else
                    {
                        // If no. of members are 5 and the required groups are 3,
                        // then we create 3 groups with 3 members + creator and add remaining members one by one in each of the groups created.
                        for (int channel = 0; channel < teamsChannels.Count; channel++)
                        {
                            channelsGroupingwithMembers[channel].Add(teamsChannels[channel]);
                        }
                    }
                }

                return channelsGroupingwithMembers;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while performing grouping logic based on given number of groups to be created.");
                return null;
            }
        }

        /// <summary>
        /// Method sends grouping of members with channels message in teams channel from where it is invoked.
        /// </summary>
        /// <param name="valuesFromTaskModule">Values enter by user in task module.</param>
        /// <param name="membersGroupingWithChannel">A dictionary of members grouped in channels.</param>
        /// <param name="groupActivityCreator">Team owner who initiated the group activity.</param>
        /// <returns>Returns a message to be shown in teams channels.</returns>
        public string GroupingMessage(GroupDetail valuesFromTaskModule, Dictionary<int, IList<TeamsChannelAccount>> membersGroupingWithChannel, string groupActivityCreator)
        {
            try
            {
                StringBuilder groupMessageActivity = new StringBuilder();
                StringBuilder membersName = new StringBuilder();

                int channelCount = 1;
                foreach (var groups in membersGroupingWithChannel)
                {
                    membersName.Append(" ");
                    string groupTextCounter = $"Group-{channelCount}";
                    string groupName = $"{valuesFromTaskModule.GroupTitle.Trim()}";

                    // limiting the text content to show till 40 characters in adaptive card
                    string truncatedGroupName = groupName.Length <= TruncateThresholdLength ? groupName : groupName.Substring(0, 40) + "...";
                    foreach (var member in groups.Value)
                    {
                        if (!member.Name.Equals(groupActivityCreator, StringComparison.OrdinalIgnoreCase))
                        {
                            membersName.Append(member.Name).Append(",").Append(" ");
                        }
                    }

                    channelCount++;
                    string groupMessageActivityText = $"**{groupTextCounter}** - **{truncatedGroupName}** :";
                    groupMessageActivity.AppendLine(groupMessageActivityText).AppendLine().AppendLine(membersName.ToString().Trim().TrimEnd(',')).AppendLine();
                    membersName.Clear();
                }

                this.logger.LogInformation($"Grouping message is : {groupMessageActivity.ToString()} ");
                return groupMessageActivity.ToString();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while creating grouping message that is to be posted after grouping.");
                throw;
            }
        }
    }
}
