// <copyright file="ChannelHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Xml;
    using Microsoft.AspNetCore.Mvc.ModelBinding.Validation;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GroupBot.Cards;
    using Microsoft.Teams.Apps.GroupBot.Common.Interfaces;
    using Microsoft.Teams.Apps.GroupBot.Models;
    using Microsoft.Teams.Apps.GroupBot.Models.Configurations;
    using Microsoft.Teams.Apps.GroupBot.Models.PrivateChannel;
    using Microsoft.Teams.Apps.GroupBot.Resources;
    using Newtonsoft.Json;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// Class to create public and private channels based on grouping criteria.
    /// </summary>
    public class ChannelHelper : IChannelHelper
    {
        /// <summary>
        /// Maximum count of private channels that are allowed to be created.
        /// </summary>
        private const int PrivateChannelCount = 30;

        /// <summary>
        /// Maximum count of public channels that are allowed to be created.
        /// </summary>
        private const int PublicChannelCount = 200;

        /// <summary>
        /// Membership type of public channel.
        /// </summary>
        private const string StandardChannelMembershipType = "standard";

        /// <summary>
        /// Membership type of private channel.
        /// </summary>
        private const string PrivateChannelMembershipType = "private";

        /// <summary>
        /// Role of user as owner in channel.
        /// </summary>
        private const string OwnerUserRole = "owner";

        /// <summary>
        /// Role of user as member in channel.
        /// </summary>
        private const string MemberUserRole = "member";

        /// <summary>
        /// Graph API for channel member while creating channels.
        /// </summary>
        private const string ChannelMemberGraphUrl = "https://graph.microsoft.com/beta/users";

        /// <summary>
        /// Retry policy with jitter, Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </summary>
        private static AsyncRetryPolicy retryPolicy = Policy.Handle<Exception>()
            .WaitAndRetryAsync(Backoff.DecorrelatedJitterBackoffV2(TimeSpan.FromMilliseconds(1000), 2));

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Helper for sending group notification in public channel.
        /// </summary>
        private readonly IGroupNotificationStorageHelper groupNotificationStorageHelper;

        /// <summary>
        /// Helper for accessing Microsoft Graph API.
        /// </summary>
        private readonly IGraphApiHelper graphApiHelper;

        /// <summary>
        /// Helper for storing group activity details into storage.
        /// </summary>
        private readonly IGroupActivityStorageHelper groupActivityStorageHelper;

        // TenantId.
        private readonly string tenantId;

        /// <summary>
        /// App credentials.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Represents a set of key/value application configuration properties for Group Activities bot.
        /// </summary>
        private readonly BotAppSetting options;

        /// <summary>
        /// Used to run a background task using IHostedService.
        /// </summary>
        private readonly BackgroundTaskWrapper taskWrapper;

        /// <summary>
        /// Initializes a new instance of the <see cref="ChannelHelper"/> class.
        /// </summary>
        /// <param name="microsoftAppCredentials">App Credentials.</param>
        /// <param name="optionsAccessor">A set of key/value application configuration properties for Group activities bot.</param>
        /// <param name="groupActivityStorageHelper">Helper for storing group activity details in azure table storage.</param>
        /// <param name="graphApiHelper">Helper for accessing Microsoft Graph API.</param>
        /// <param name="groupNotificationStorageHelper">Helper for sending notification in channel.</param>
        /// <param name="backgroundTaskWrapper">Instance for background task wrapper.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public ChannelHelper(ILogger<ChannelHelper> logger, IOptionsMonitor<BotAppSetting> optionsAccessor, IGroupNotificationStorageHelper groupNotificationStorageHelper, IGroupActivityStorageHelper groupActivityStorageHelper, IGraphApiHelper graphApiHelper, MicrosoftAppCredentials microsoftAppCredentials, BackgroundTaskWrapper backgroundTaskWrapper)
        {
            this.options = optionsAccessor.CurrentValue;
            this.graphApiHelper = graphApiHelper;
            this.logger = logger;
            this.groupNotificationStorageHelper = groupNotificationStorageHelper;
            this.groupActivityStorageHelper = groupActivityStorageHelper;
            this.tenantId = this.options.TenantId;
            this.microsoftAppCredentials = microsoftAppCredentials;
            this.taskWrapper = backgroundTaskWrapper;
        }

        /// <summary>
        /// Methods stores newly created group activity into azure table storage.
        /// </summary>
        /// <param name="serviceUrl">Bot activity service URL.</param>
        /// <param name="groupActivityId">Group activity Id.</param>
        /// <param name="teamId">Team id of team where bot is installed.</param>
        /// <param name="groupDetail">Group activity details entered by users in task module.</param>
        /// <param name="groupActivityCreator">Team owner who initiated the group activity.</param>
        /// <param name="groupingCardConversationId">Conversation id of group detail card sent into channel.</param>
        /// <param name="groupingCardActivityId">Activity id of group detail card sent into channel.</param>
        /// <returns>A task that stores group activity data to table storage.</returns>
        public async Task StoreGroupActivityDetailsAsync(string serviceUrl, string groupActivityId, string teamId, GroupDetail groupDetail, string groupActivityCreator, string groupingCardConversationId, string groupingCardActivityId)
        {
            try
            {
                GroupActivityEntity groupActivityEntity = new GroupActivityEntity
                {
                    GroupActivityId = groupActivityId,
                    GroupActivityTitle = groupDetail.GroupTitle,
                    GroupActivityDescription = groupDetail.GroupDescription,
                    IsPrivateChannel = groupDetail.ChannelType == Constants.PrivateChannelType ? true : false,
                    DueDate = groupDetail.DueDate,
                    CreatedOn = DateTime.UtcNow,
                    CreatedBy = groupActivityCreator,
                    ConversationId = groupingCardConversationId,
                    ActivityId = groupingCardActivityId,
                    TeamId = teamId,
                    ServiceUrl = serviceUrl,
                    IsNotificationActive = groupDetail.ChannelType == Constants.PrivateChannelType ? false : (groupDetail.AutoReminders == Constants.AutoReminderNo ? false : true),
                };

                var isGroupActivitySaved = await this.groupActivityStorageHelper.UpsertGroupActivityAsync(groupActivityEntity);
                if (!isGroupActivitySaved)
                {
                    this.logger.LogError($"Error while saving Group activity data in storage for teamId : {teamId}");
                    return;
                }

                this.logger.LogInformation($"Group activity data successfully saved in storage for teamId : {teamId}");
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while saving Group activity data in storage for teamId : {teamId}");
                throw;
            }
        }

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
        public async Task ValidateAndCreateChannelAsync(string token, string groupActivityId, string teamId, string groupId, string groupingMessage, GroupDetail valuesFromTaskModule, Dictionary<int, IList<TeamsChannelAccount>> membersGroupingWithChannel, TeamsChannelAccount groupActivityCreator, ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            // Validate that channel count to be created are in limit with public channels less than 200 and private channel 30.
            bool? isValidChannelCount = await this.ValidateChannelCountAsync(token, valuesFromTaskModule.ChannelType, membersGroupingWithChannel.Count, groupId);

            if (isValidChannelCount == null)
            {
                await turnContext.SendActivityAsync(Strings.CustomErrorMessage);
                this.logger.LogInformation($"Channel count is null for teamId : {teamId}");
                return;
            }

            if (isValidChannelCount == false)
            {
                await turnContext.SendActivityAsync(Strings.ChannelCountValidationText);
                this.logger.LogInformation($"Channel count is not valid for teamId : {teamId}");
                return;
            }

            this.taskWrapper.Enqueue(this.CreateChannelAsync(token, groupActivityId, teamId, groupId, valuesFromTaskModule, membersGroupingWithChannel, groupActivityCreator, groupingMessage, turnContext, cancellationToken));
        }

        /// <summary>
        /// Create public or private channel based on grouping criteria.
        /// </summary>
        /// <param name="accessToken">Token to access Microsoft Graph API.</param>
        /// <param name="groupActivityId">group activity Id.</param>
        /// <param name="teamId">Team id where messaging extension is installed.</param>
        /// <param name="groupId">Team Azure Active Directory object id of the channel where bot is installed.</param>
        /// <param name="valuesFromTaskModule">Group activity details from task module entered by user.</param>
        /// <param name="membersGroupingWithChannel">List of all members grouped in channels based on grouping criteria.</param>
        /// <param name="groupActivityCreator">Team owner who started the group activity.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that returns true if channel is created successfully.</returns>
        private async Task CreateChannelAsync(string accessToken, string groupActivityId, string teamId, string groupId, GroupDetail valuesFromTaskModule, Dictionary<int, IList<TeamsChannelAccount>> membersGroupingWithChannel, TeamsChannelAccount groupActivityCreator, string groupingMessage, ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                string channelType = valuesFromTaskModule.ChannelType;

                Tuple<List<ChannelApiResponse>, List<string>> channelDetails = null;
                switch (channelType)
                {
                    case Constants.PublicChannelType:
                        channelDetails = await this.CreatePublicChannelAsync(accessToken, membersGroupingWithChannel, valuesFromTaskModule, groupId, groupActivityCreator.Name, groupingMessage, turnContext, cancellationToken);
                        break;
                    case Constants.PrivateChannelType:
                        channelDetails = await this.CreatePrivateChannelAsync(accessToken, groupId, membersGroupingWithChannel, groupActivityCreator, valuesFromTaskModule);
                        break;
                }

                if (channelDetails != null && channelDetails.Item1.Count > 0)
                {
                    bool isChannelInfoSaved = await this.StoreChannelsCreatedDetailsAsync(teamId, groupActivityId, channelType, channelDetails.Item1, turnContext);
                    if (!isChannelInfoSaved)
                    {
                        await turnContext.SendActivityAsync(Strings.CustomErrorMessage);
                        this.logger.LogInformation($"Saving newly created channel details to table storage failed for teamId: {teamId}.");
                    }
                }

                // show the list of channels to the user which are failed to get created.
                if (channelDetails.Item2.Count > 0)
                {
                    StringBuilder channelsNotCreated = new StringBuilder();
                    foreach (var channel in channelDetails.Item2)
                    {
                        channelsNotCreated.AppendLine(channel).AppendLine();
                    }

                    this.logger.LogError($"Number of channels failed to get created. TotalChannels {channelDetails.Item2.Count.ToString()}, Team {teamId}");
                    await turnContext.SendActivityAsync(MessageFactory.Attachment(GroupActivityCard.GetChannelCreationFailedCard(channelsNotCreated.ToString(), valuesFromTaskModule.GroupTitle)));
                }
            }
            catch (Exception ex)
            {
                await turnContext.SendActivityAsync(Strings.CustomErrorMessage);
                this.logger.LogError(ex, $"Error while creating channels for teamId: {teamId}");
            }
        }

        /// <summary>
        /// Method to validate whether channels count to be created are in limit with public channels less than 200 and private channel 30.
        /// </summary>
        /// <param name="accessToken">Token to access Microsoft Graph API.</param>
        /// <param name="channelType">Channel type that is public or private.</param>
        /// <param name="noOfChannelsToCreate">Count of number of channels to be created.</param>
        /// <param name="groupId">Team Azure Active Directory object id of the channel where bot is installed.</param>
        /// <returns>A task that returns true if channel count is valid.</returns>
        private async Task<bool?> ValidateChannelCountAsync(string accessToken, string channelType, int noOfChannelsToCreate, string groupId)
        {
            try
            {
                bool isValidChannelCount = true;
                var getChannelsList = await this.graphApiHelper.GetChannelsAsync(accessToken, groupId);

                if (getChannelsList != null && channelType.Equals(Constants.PublicChannelType))
                {
                    var publicChannelCount = getChannelsList.ChannelsValue.Count(x => x.MembershipType.Equals(StandardChannelMembershipType));
                    isValidChannelCount = PublicChannelCount - publicChannelCount > noOfChannelsToCreate;
                }
                else if (getChannelsList != null && channelType.Equals(Constants.PrivateChannelType))
                {
                    var privateChannelsCount = getChannelsList.ChannelsValue.Count(x => x.MembershipType.Equals(PrivateChannelMembershipType));
                    isValidChannelCount = PrivateChannelCount - privateChannelsCount > noOfChannelsToCreate;
                }

                return isValidChannelCount;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while validating count of channels.");
                return null;
            }
        }

        /// <summary>
        /// Store all newly created channel details for notifications in azure table storage.
        /// </summary>
        /// <param name="teamId">team Id of the team in which bot is installed.</param>
        /// <param name="groupActivityId">Guid group activity Id.</param>
        /// <param name="channelType">Channel type that is private or public.</param>
        /// <param name="publicChannelApiResponses">Api response obtained by creating channels.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <returns>A task that is true if data is successfully saved.</returns>
        private async Task<bool> StoreChannelsCreatedDetailsAsync(string teamId, string groupActivityId, string channelType, List<ChannelApiResponse> publicChannelApiResponses, ITurnContext<IInvokeActivity> turnContext)
        {
            string conversationId = turnContext.Activity.Conversation.Id;
            string serviceUrl = turnContext.Activity.ServiceUrl;
            List<GroupNotification> groupNotificationsList = new List<GroupNotification>();
            bool isNotificationActive = channelType.Equals(Constants.PublicChannelType) ? true : false;

            try
            {
                groupNotificationsList.AddRange(publicChannelApiResponses.Select(channel => new GroupNotification()
                {
                    TeamId = teamId,
                    ChannelId = channel.Id,
                    ChannelName = channel.DisplayName,
                    GroupActivityId = groupActivityId,
                }));

                return await this.groupNotificationStorageHelper.UpsertGroupNotificationDetailsBatchAsync(groupNotificationsList);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while saving created channels details to table storage for notification for teamId: {teamId}.");
                return false;
            }
        }

        /// <summary>
        /// Method creates a public channel based on grouping result obtained from grouping criteria.
        /// </summary>
        /// <param name="accessToken">Token to access Microsoft Graph API.</param>
        /// <param name="membersGroupingWithChannel">A dictionary with members grouped into channels based on entered grouping criteria.</param>
        /// <param name="groupDetail">Values obtained from task modules which entered by user.</param>
        /// <param name="groupId">Team Azure Active Directory object id of the channel where bot is installed.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that has list of public channel created.</returns>
        private async Task<Tuple<List<ChannelApiResponse>, List<string>>> CreatePublicChannelAsync(string accessToken, Dictionary<int, IList<TeamsChannelAccount>> membersGroupingWithChannel, GroupDetail groupDetail, string groupId, string groupActivityCreator, string groupingMessage, ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            List<ChannelApiResponse> publicChannelApiResponses = new List<ChannelApiResponse>();
            var notCreatedChannels = new List<string>();
            try
            {
                var createChannelRequest = new PublicChannelRequest();
                int channelCount = 1;
                foreach (var groupedChannel in membersGroupingWithChannel)
                {
                    createChannelRequest.DisplayName = $"{groupDetail.GroupTitle}-{channelCount}";
                    createChannelRequest.Description = groupDetail.GroupDescription;

                    string createChannelData = JsonConvert.SerializeObject(createChannelRequest);

                    var channelCreated = await this.graphApiHelper.CreatePublicChannelAsync(accessToken, createChannelData, groupId);

                    if (channelCreated != null)
                    {
                        this.logger.LogInformation($"Successfully created public channel : {channelCreated.DisplayName} ");

                        // Mention members in the channels where they are grouped.
                        await retryPolicy.ExecuteAsync(async () =>
                        {
                            await this.NotifyGroupMembersInChannel(groupedChannel.Value, channelCreated.Id, groupDetail, groupActivityCreator, groupingMessage, turnContext, cancellationToken);
                        });

                        publicChannelApiResponses.Add(channelCreated);
                    }
                    else
                    {
                        notCreatedChannels.Add(string.Format("*{0}*", createChannelRequest.DisplayName));
                        this.logger.LogInformation($"Not able to create channel for: {createChannelRequest.DisplayName} ");
                    }

                    channelCount++;
                }

                return new Tuple<List<ChannelApiResponse>, List<string>>(publicChannelApiResponses, notCreatedChannels);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while creating public channel for groupId : {groupId}");
                return null;
            }
        }

        private async Task NotifyGroupMembersInChannel(IList<TeamsChannelAccount> groupMembers, string channelId, GroupDetail groupDetail, string groupActivityCreator, string groupingMessage, ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            var activity = turnContext.Activity as Activity;
            var channelData = activity.GetChannelData<TeamsChannelData>();
            channelData.Channel.Id = channelId;

            // create card
            var notificationCard = MessageFactory.Attachment(NotificationCard.GetNewMemberNotificationCard(groupDetail, groupActivityCreator));

            // get mentions
            var mentionActivity = this.GetGroupMembersToMentionActivity(groupMembers);

            var conversationParameters = new ConversationParameters
            {
                Activity = (Activity)notificationCard,
                Bot = new ChannelAccount { Id = this.microsoftAppCredentials.MicrosoftAppId },
                ChannelData = channelData,
                TenantId = channelData.Tenant.Id,
            };

            await ((BotFrameworkAdapter)turnContext.Adapter).CreateConversationAsync(
                Constants.Channel,
                turnContext.Activity.ServiceUrl,
                this.microsoftAppCredentials,
                conversationParameters,
                async (newTurnContext, newCancellationToken) =>
                {
                    await turnContext.Adapter.ContinueConversationAsync(
                        this.microsoftAppCredentials.MicrosoftAppId,
                        newTurnContext.Activity.GetConversationReference(),
                        async (conversationTurnContext, conversationCancellationToken) =>
                        {
                            mentionActivity.ApplyConversationReference(conversationTurnContext.Activity.GetConversationReference());
                            await conversationTurnContext.SendActivityAsync(mentionActivity, conversationCancellationToken);
                        },
                        newCancellationToken);
                },
                cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Method creates a private channel based on grouping result obtained from grouping criteria.
        /// </summary>
        /// <param name="accessToken">Token to access Microsoft Graph API.</param>
        /// <param name="groupId">Team Azure Active Directory object id of the channel where bot is installed.</param>
        /// <param name="membersGroupingWithChannel">A dictionary with members grouped into channels based on entered grouping criteria.</param>
        /// <param name="groupActivityCreator">Team owner who initiated group activity.</param>
        /// <param name="groupDetail">Values entered by user in task module.</param>
        /// <returns>Return the List of channels created using Microsoft Graph API.</returns>
        private async Task<Tuple<List<ChannelApiResponse>, List<string>>> CreatePrivateChannelAsync(string accessToken, string groupId, Dictionary<int, IList<TeamsChannelAccount>> membersGroupingWithChannel, TeamsChannelAccount groupActivityCreator, GroupDetail groupDetail)
        {
            List<ChannelApiResponse> privateChannelApiResponses = new List<ChannelApiResponse>();
            var notCreatedChannels = new List<string>();
            try
            {
                var privateChannelRequestData = new PrivateChannelRequest();
                var channelCount = 1;
                foreach (var groupedChannel in membersGroupingWithChannel)
                {
                    privateChannelRequestData.DisplayName = $"{groupDetail.GroupTitle}-{channelCount}";
                    privateChannelRequestData.Description = groupDetail.GroupDescription;
                    privateChannelRequestData.MembershipType = groupDetail.ChannelType;
                    privateChannelRequestData.Members = new List<ChannelMember>();

                    foreach (var groupMember in groupedChannel.Value)
                    {
                        var role = groupActivityCreator.AadObjectId.Contains(groupMember.AadObjectId) ? OwnerUserRole : MemberUserRole;
                        privateChannelRequestData.Members.Add(new ChannelMember { UserOdataBind = $"{ChannelMemberGraphUrl}('{groupMember.AadObjectId}')", Roles = new List<string> { role } });
                    }

                    string createChannelData = JsonConvert.SerializeObject(privateChannelRequestData);

                    var privateChannels = await this.graphApiHelper.CreatePrivateChannelAsync(accessToken, createChannelData, groupId);
                    if (privateChannels != null)
                    {
                        privateChannelApiResponses.Add(privateChannels);
                    }
                    else
                    {
                        notCreatedChannels.Add(string.Format("*{0}*", privateChannels.DisplayName));
                        this.logger.LogInformation($"Not able to create channel for: {privateChannels.DisplayName}");
                    }

                    channelCount++;
                }

                return new Tuple<List<ChannelApiResponse>, List<string>>(privateChannelApiResponses, notCreatedChannels);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while creating private channel for the team.");
                return null;
            }
        }

        /// <summary>
        /// Method mentions user in respective channel of which they are part of after grouping.
        /// </summary>
        /// <param name="membersGroupingWithChannel">A list with members grouped into channels based on entered grouping criteria.</param>
        /// <returns>Members to mention activity.</returns>
        private Activity GetGroupMembersToMentionActivity(IList<TeamsChannelAccount> membersGroupingWithChannel)
        {
            try
            {
                StringBuilder membersMention = new StringBuilder();
                var entities = new List<Entity>();
                var mentions = new List<Mention>();

                foreach (var member in membersGroupingWithChannel)
                {
                    membersMention.Append(" ");
                    var mention = new Mention
                    {
                        Mentioned = new ChannelAccount()
                        {
                            Id = member.Id,
                            Name = member.Name,
                        },
                        Text = $"<at>{XmlConvert.EncodeName(member.Name)}</at>",
                    };
                    mentions.Add(mention);
                    entities.Add(mention);
                    membersMention.Append(mention.Text).Append(",").Append(" ");
                }

                var memberActivity = MessageFactory.Text(membersMention.ToString().Trim().TrimEnd(','));
                memberActivity.Entities = entities;

                return memberActivity;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while creating channel members to mention in channel.");
                return null;
            }
        }
    }
}
