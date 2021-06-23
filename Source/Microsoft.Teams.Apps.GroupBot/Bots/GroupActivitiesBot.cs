// <copyright file="GroupActivitiesBot.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Bots
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Net;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GroupBot.Cards;
    using Microsoft.Teams.Apps.GroupBot.Common;
    using Microsoft.Teams.Apps.GroupBot.Common.Interfaces;
    using Microsoft.Teams.Apps.GroupBot.Models;
    using Microsoft.Teams.Apps.GroupBot.Models.Configurations;
    using Microsoft.Teams.Apps.GroupBot.Resources;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Class for teams activity of Group activities bot and messaging extension.
    /// </summary>
    public class GroupActivitiesBot : TeamsActivityHandler
    {
        /// <summary>
        /// Sets the height of the task module.
        /// </summary>
        private const int TaskModuleHeight = 600;

        /// <summary>
        /// Sets the width of the task module.
        /// </summary>
        private const int TaskModuleWidth = 600;

        /// <summary>
        /// Sets the height of task module of validation message.
        /// </summary>
        private const string TaskModuleValidationHeight = "small";

        /// <summary>
        /// Sets the team members cache key.
        /// </summary>
        private const string TeamMembersCacheKey = "teamMembersCacheKey";

        /// <summary>
        /// Sets the width of task module of validation message.
        /// </summary>
        private const string TaskModuleValidationWidth = "medium";

        /// <summary>
        /// Messaging extension authentication type.
        /// </summary>
        private const string MessagingExtensionAuthType = "auth";

        /// <summary>
        /// Messaging extension default parameter value.
        /// </summary>
        private const string MessagingExtensionInitialParameterName = "initialRun";

        /// <summary>
        /// Messaging extension result type.
        /// </summary>
        private const string MessagingExtenstionResultType = "result";

        /// <summary>
        /// Messaging extension message type.
        /// </summary>
        private const string MessagingExtensionMessageType = "message";

        /// <summary>
        /// Command Id for recent activities in messaging extension.
        /// </summary>
        private const string RecentCommandId = "recentActivities";

        /// <summary>
        /// Command Id for all activities in messaging extension.
        /// </summary>
        private const string AllCommandId = "allActivities";

        /// <summary>
        /// Helper for creating channels for grouped members.
        /// </summary>
        private readonly IChannelHelper channelHelper;

        /// <summary>
        /// Helper for storing group activity details into storage.
        /// </summary>
        private readonly IGroupActivityStorageHelper groupActivityStorageHelper;

        /// <summary>
        /// Helper for grouping members based on splitting criteria.
        /// </summary>
        private readonly IGroupingHelper groupingHelper;

        /// <summary>
        /// Helper to get group members and verify if user is a team owner.
        /// </summary>
        private readonly ITeamUserHelper teamUserHelper;

        /// <summary>
        /// Azure Active Directory (AADv2) bot connection name.
        /// </summary>
        private readonly string connectionName;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Represents a set of key/value application configuration properties for Group activities bot.
        /// </summary>
        private readonly BotAppSetting options;

        /// <summary>
        /// Application base URI.
        /// </summary>
        private readonly string appBaseUrl;

        /// <summary>
        /// Tenant id.
        /// </summary>
        private readonly string tenantId;

        /// <summary>
        /// Telemetry client to log events.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="GroupActivitiesBot"/> class.
        /// </summary>
        /// <param name="optionsAccessor">A set of key/value application configuration properties for Group activities bot.</param>
        /// <param name="channelHelper">Helper for creating public and private channels.</param>
        /// <param name="groupingHelper">Helper for grouping members to channels based on specified splitting criteria.</param>
        /// <param name="groupActivityStorageHelper">Helper for storing group activity in azure table storage.</param>
        /// <param name="teamUserHelper">Helper to get group members and verify if user is a team owner.</param>
        /// <param name="telemetryClient">Telemetry client.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public GroupActivitiesBot(IOptionsMonitor<BotAppSetting> optionsAccessor, IChannelHelper channelHelper, IGroupingHelper groupingHelper, IGroupActivityStorageHelper groupActivityStorageHelper, ITeamUserHelper teamUserHelper, TelemetryClient telemetryClient, ILogger<GroupActivitiesBot> logger, IMemoryCache memoryCache)
        {
            this.options = optionsAccessor.CurrentValue;
            this.tenantId = this.options.TenantId;
            this.appBaseUrl = this.options.AppBaseURI;
            this.connectionName = this.options.ConnectionName;
            this.channelHelper = channelHelper;
            this.groupingHelper = groupingHelper;
            this.groupActivityStorageHelper = groupActivityStorageHelper;
            this.teamUserHelper = teamUserHelper;
            this.telemetryClient = telemetryClient;
            this.logger = logger;
        }

        /// <summary>
        /// Method will be invoked on each bot turn.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public override async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default)
        {
            var activity = turnContext.Activity;
            if (!this.IsActivityFromExpectedTenant(turnContext))
            {
                this.logger.LogInformation($"Unexpected tenant Id {activity.Conversation.TenantId}", SeverityLevel.Warning);
                await turnContext.SendActivityAsync(activity: MessageFactory.Text(Strings.InvalidTenant));
            }
            else
            {
                // Get the current culture info to use in resource files
                string locale = activity.Entities?.Where(entity => entity.Type == "clientInfo").First().Properties["locale"].ToString();

                if (!string.IsNullOrEmpty(locale))
                {
                    CultureInfo.CurrentCulture = CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo(locale);
                }

                await base.OnTurnAsync(turnContext, cancellationToken);
            }
        }

        /// <summary>
        /// Handle message extension action fetch task received by the bot.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="action">Messaging extension action value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that returns messagingExtensionActionResponse.</returns>
        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionFetchTaskAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            var activity = turnContext.Activity;

            var activityState = ((JObject)activity.Value).GetValue("state")?.ToString();
            var tokenResponse = await (turnContext.Adapter as IUserTokenProvider).GetUserTokenAsync(turnContext, this.connectionName, activityState, cancellationToken);

            if (tokenResponse == null)
            {
                var signInLink = await (turnContext.Adapter as IUserTokenProvider).GetOauthSignInLinkAsync(turnContext, this.connectionName, cancellationToken);

                return new MessagingExtensionActionResponse
                {
                    ComposeExtension = new MessagingExtensionResult
                    {
                        Type = MessagingExtensionAuthType,
                        SuggestedActions = new MessagingExtensionSuggestedAction
                        {
                            Actions = new List<CardAction>
                                {
                                    new CardAction
                                    {
                                        Type = ActionTypes.OpenUrl,
                                        Value = signInLink,
                                        Title = Strings.SigninCardText,
                                    },
                                },
                        },
                    },
                };
            }

            var teamInformation = activity.TeamsGetTeamInfo();
            if (teamInformation == null || string.IsNullOrEmpty(teamInformation.Id))
            {
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Card = GroupActivityCard.GetTeamNotFoundErrorCard(),
                            Height = TaskModuleValidationHeight,
                            Width = TaskModuleValidationWidth,
                            Title = Strings.GroupActivityTitle,
                        },
                    },
                };
            }

            TeamDetails teamDetails;
            try
            {
                teamDetails = await TeamsInfo.GetTeamDetailsAsync(turnContext, teamInformation.Id, cancellationToken);
            }
            catch (Exception ex)
            {
                // if bot is not installed in team or not able to team roster, then show error response.
                this.logger.LogError("Bot is not part of team roster", ex);
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo()
                        {
                            Card = GroupActivityCard.GetTeamNotFoundErrorCard(),
                            Height = TaskModuleHeight,
                            Width = TaskModuleHeight,
                        },
                    },
                };
            }

            if (teamDetails == null)
            {
                this.logger.LogInformation($"Team details obtained is null.");
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Card = GroupActivityCard.GetErrorMessageCard(),
                            Height = TaskModuleValidationHeight,
                            Width = TaskModuleValidationWidth,
                            Title = Strings.GroupActivityTitle,
                        },
                    },
                };
            }

            var isTeamOwner = await this.teamUserHelper.VerifyIfUserIsTeamOwnerAsync(tokenResponse.Token, teamDetails.AadGroupId, activity.From.AadObjectId);
            if (isTeamOwner == null)
            {
                await turnContext.SendActivityAsync(Strings.CustomErrorMessage);
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Card = GroupActivityCard.GetErrorMessageCard(),
                            Height = TaskModuleValidationHeight,
                            Width = TaskModuleValidationWidth,
                            Title = Strings.GroupActivityTitle,
                        },
                    },
                };
            }

            // If user is team member validation message is shown as only team owner can create a group activity.
            if (isTeamOwner == false)
            {
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Card = GroupActivityCard.GetTeamOwnerErrorCard(),
                            Height = TaskModuleValidationHeight,
                            Width = TaskModuleValidationWidth,
                            Title = Strings.GroupActivityTitle,
                        },
                    },
                };
            }

            // Team owner can create group activity.
            return new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        Card = GroupActivityCard.GetCreateGroupActivityCard(),
                        Height = TaskModuleHeight,
                        Width = TaskModuleWidth,
                        Title = Strings.GroupActivityTitle,
                    },
                },
            };
        }

        /// <summary>
        /// When OnTurn method receives a compose extension query invoke activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="query">Messaging extension query request value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents messaging extension response.</returns>
        protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
        {
            try
            {
                // Execute code when parameter name is initial run.
                if (query.Parameters.First().Name == MessagingExtensionInitialParameterName)
                {
                    this.logger.LogInformation("Executing initial run parameter from messaging extension.");

                    // Get access token for user.if already authenticated, we will get token.
                    // If user is not signed in, send sign in link in messaging extension.
                    var tokenResponse = await (turnContext.Adapter as IUserTokenProvider).GetUserTokenAsync(turnContext, this.connectionName, query.State, cancellationToken);

                    if (tokenResponse == null)
                    {
                        var signInLink = await (turnContext.Adapter as IUserTokenProvider).GetOauthSignInLinkAsync(turnContext, this.connectionName, cancellationToken);
                        return new MessagingExtensionResponse
                        {
                            ComposeExtension = new MessagingExtensionResult
                            {
                                Type = MessagingExtensionAuthType,
                                SuggestedActions = new MessagingExtensionSuggestedAction
                                {
                                    Actions = new List<CardAction>
                                    {
                                        new CardAction
                                        {
                                            Type = ActionTypes.OpenUrl,
                                            Value = signInLink,
                                            Title = Strings.SigninCardText,
                                        },
                                    },
                                },
                            },
                        };
                    }
                }

                return await this.HandleMessagingExtensionSearchQueryAsync(turnContext, query);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error in handling invoke action from messaging extension.");
                return null;
            }
        }

        /// <summary>
        /// Handle message extension submit action received by the bot.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="action">Messaging extension action request value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents messaging extension response.</returns>
        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            var activity = turnContext.Activity;
            Dictionary<int, IList<TeamsChannelAccount>> membersGroupingWithChannel = null;

            try
            {
                var valuesFromTaskModule = JsonConvert.DeserializeObject<GroupDetail>(action.Data?.ToString());
                var teamInformation = activity.TeamsGetTeamInfo();

                TokenResponse tokenResponse = await (turnContext.Adapter as IUserTokenProvider).GetUserTokenAsync(turnContext, this.connectionName, null, cancellationToken);
                string token = tokenResponse.Token;

                if (token == null || valuesFromTaskModule == null)
                {
                    this.logger.LogInformation($"Either token obtained is null. Token : {token} or values obtained from task module is null for {teamInformation.Id}");
                    await turnContext.SendActivityAsync(Strings.CustomErrorMessage);
                    return new MessagingExtensionActionResponse();
                }

                // Activity local timestamp provides offset value which can be used to convert user input time with offset.
                valuesFromTaskModule.DueDate = new DateTimeOffset(
                    valuesFromTaskModule.DueDate.Year,
                    valuesFromTaskModule.DueDate.Month,
                    valuesFromTaskModule.DueDate.Day,
                    valuesFromTaskModule.DueTime.Hour,
                    valuesFromTaskModule.DueTime.Minute,
                    valuesFromTaskModule.DueTime.Second,
                    turnContext.Activity.LocalTimestamp.Value.Offset).ToUniversalTime();

                // Validate task module values entered by user to create group activity.
                if (!this.IsValidGroupActivityInputFields(valuesFromTaskModule))
                {
                    return this.ShowValidationForCreateChannel(valuesFromTaskModule, isSplittingValid: true);
                }

                var teamMembers = new List<TeamsChannelAccount>();
                string continuationToken = null;

                do
                {
                    var currentPage = await TeamsInfo.GetPagedTeamMembersAsync(turnContext, teamInformation.Id, continuationToken, pageSize: 500);
                    continuationToken = currentPage.ContinuationToken;
                    teamMembers.AddRange(currentPage.Members);
                }
                while (continuationToken != null);

                if (teamMembers == null)
                {
                    this.logger.LogError($"List of channel members and owners in a team obtained is null for team id : {teamInformation.Id}");
                    return new MessagingExtensionActionResponse();
                }

                var teamDetails = await TeamsInfo.GetTeamDetailsAsync(turnContext, teamInformation.Id, cancellationToken);

                // Get list of members to perform grouping excluding team owners of the team.
                var groupMembers = await this.teamUserHelper.GetGroupMembersAsync(token, teamDetails.AadGroupId, teamMembers);
                if (groupMembers.Count() <= 0 || groupMembers == null)
                {
                    await turnContext.SendActivityAsync(Strings.TeamMembersDoesNotExistsText);
                    this.logger.LogInformation($"Group members obtained to perform grouping is null for team id : {teamInformation.Id}.");
                    return new MessagingExtensionActionResponse();
                }

                // Identify the team owner who initiated the group activity.
                var groupActivityCreator = teamMembers.Where(members => members.AadObjectId.Contains(activity.From.AadObjectId)).FirstOrDefault();

                // Group members based on splitting criteria entered by user.
                switch (valuesFromTaskModule.SplittingCriteria)
                {
                    case Constants.SplitInGroupOfGivenMembers:
                        membersGroupingWithChannel = this.groupingHelper.SplitInGroupOfGivenMembers(groupMembers, groupActivityCreator, valuesFromTaskModule.ChannelOrMemberUnits);
                        break;

                    case Constants.SplitInGivenNumberOfGroups:
                        int numberOfMembersInEachGroup = groupMembers.Count() / valuesFromTaskModule.ChannelOrMemberUnits;

                        // Validation to not allow for grouping if member in each group is 1.
                        if (groupMembers.Count() <= valuesFromTaskModule.ChannelOrMemberUnits)
                        {
                            return this.ShowValidationForCreateChannel(valuesFromTaskModule, isSplittingValid: false);
                        }

                        membersGroupingWithChannel = this.groupingHelper.SplitInGivenNumberOfGroups(groupMembers, groupActivityCreator, valuesFromTaskModule.ChannelOrMemberUnits, numberOfMembersInEachGroup);
                        break;
                }

                if (membersGroupingWithChannel == null || membersGroupingWithChannel.Count <= 0)
                {
                    this.logger.LogError($"Error while grouping members to channel: {activity.Conversation.Id}");
                    return new MessagingExtensionActionResponse();
                }

                string groupActivityId = Guid.NewGuid().ToString();
                string groupingMessage = await this.SendAndStoreGroupingMessageAsync(groupActivityId, teamDetails.Id, valuesFromTaskModule, membersGroupingWithChannel, groupActivityCreator, turnContext, cancellationToken);

                // If field auto create channel is yes then create channels and send grouping message else send only grouping message in channel.
                if (valuesFromTaskModule.AutoCreateChannel == Constants.AutoCreateChannelYes)
                {
                    await this.channelHelper.ValidateAndCreateChannelAsync(token, groupActivityId, teamInformation.Id, teamDetails.AadGroupId, groupingMessage, valuesFromTaskModule, membersGroupingWithChannel, groupActivityCreator, turnContext, cancellationToken);
                }

                // Logs Click through on activity created.
                this.telemetryClient.TrackEvent("Group activity created", new Dictionary<string, string>() { { "Team", teamInformation.Id }, { "AADObjectId", activity.From.AadObjectId } });
                return new MessagingExtensionActionResponse();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while submitting data from task module through messaging extension action.");
                return null;
            }
        }

        /// <summary>
        /// Send Welcome card in teams channel.
        /// </summary>
        /// <param name="membersAdded">A list of all the members added to the conversation, as described by the conversation update activity.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Welcome card  when Bot/Messaging Extension is added first time by user.</returns>
        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var activity = turnContext.Activity;
            this.logger.LogInformation($"conversationType: {activity.Conversation.ConversationType}, membersAdded: {activity.MembersAdded?.Count}, membersRemoved: {activity.MembersRemoved?.Count}");

            // If added member is bot then send welcome card in channel.
            if (membersAdded.Any(member => member.Id == activity.Recipient.Id))
            {
                this.logger.LogInformation($"Bot added {activity.Conversation.Id}");
                await turnContext.SendActivityAsync(MessageFactory.Attachment(WelcomeCard.GetWelcomeCardAttachment(this.appBaseUrl)));
            }
        }

        /// <summary>
        /// Verify if the tenant Id in the message is the same tenant Id used when application was configured.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <returns>True if context is from expected tenant else false.</returns>
        private bool IsActivityFromExpectedTenant(ITurnContext turnContext)
        {
            return turnContext.Activity.Conversation.TenantId == this.tenantId;
        }

        /// <summary>
        /// Handles messaging extension user search query request.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="query">Messaging extension query request value payload.</param>
        /// <returns>A task that represents messaging extension response containing all group activities created by owners of the team.</returns>
        private async Task<MessagingExtensionResponse> HandleMessagingExtensionSearchQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query)
        {
            try
            {
                var activity = turnContext.Activity;
                var teamInformation = activity.TeamsGetTeamInfo();

                if (teamInformation == null)
                {
                    return new MessagingExtensionResponse
                    {
                        ComposeExtension = new MessagingExtensionResult()
                        {
                            Type = MessagingExtensionMessageType,
                            Text = Strings.NoTeamFoundErrorText,
                        },
                    };
                }

                var searchQuery = query.Parameters.First().Value.ToString();

                this.logger.LogInformation($"searchQuery : {searchQuery} commandId : {query.CommandId}");

                return new MessagingExtensionResponse
                {
                    ComposeExtension = await this.GetSearchResultAsync(teamInformation.Id, searchQuery, query.CommandId, query.QueryOptions.Count ?? 0, query.QueryOptions.Skip ?? 0).ConfigureAwait(false),
                };
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Exception while handling messaging extension query for {query.CommandId}");
                return null;
            }
        }

        /// <summary>
        /// Get messaging extension search result based on user search query and command.
        /// </summary>
        /// <param name="teamId">Team Id where bot is installed.</param>
        /// <param name="searchQuery">User search query text.</param>
        /// <param name="commandId">Messaging extension command id e.g. recentActivities or allActivities.</param>
        /// <param name="count">Count for pagination in Messaging extension.</param>
        /// <param name="skip">Skip for pagination in Messaging extension.</param>
        /// <returns>A task that represents compose extension result with activities.</returns>
        private async Task<MessagingExtensionResult> GetSearchResultAsync(string teamId, string searchQuery, string commandId, int count, int skip)
        {
            try
            {
                MessagingExtensionResult composeExtensionResult = new MessagingExtensionResult();

                IList<GroupActivityEntity> groupActivities = await this.groupActivityStorageHelper.GetGroupActivityEntityDetailAsync(teamId);

                // On initial run searchQuery value is "true".
                if (searchQuery != "true")
                {
                    this.logger.LogInformation($"search query entered by user is {searchQuery} for {teamId}");
                    groupActivities = groupActivities.Where(x => x.GroupActivityTitle.Contains(searchQuery, StringComparison.OrdinalIgnoreCase)).ToList();
                }

                // If no activities found for a team.
                if (groupActivities.Count() <= 0)
                {
                    composeExtensionResult.Type = MessagingExtensionMessageType;
                    composeExtensionResult.Text = Strings.NoActivitiesFoundText;
                    this.logger.LogInformation($"No activities found for {teamId}");
                    return composeExtensionResult;
                }

                composeExtensionResult.Type = MessagingExtenstionResultType;
                composeExtensionResult.AttachmentLayout = AttachmentLayoutTypes.List;
                switch (commandId)
                {
                    case RecentCommandId:
                        composeExtensionResult.Attachments = MessagingExtensionCard.GetGroupActivityCard(groupActivities.OrderByDescending(groupActivity => groupActivity.Timestamp).Skip(skip).Take(count).ToList());
                        break;

                    case AllCommandId:
                        composeExtensionResult.Attachments = MessagingExtensionCard.GetGroupActivityCard(groupActivities.Skip(skip).Take(count).ToList());
                        break;
                }

                return composeExtensionResult;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while generating result for messaging extension for {commandId} for {teamId}");
                return null;
            }
        }

        /// <summary>
        /// Show validation card if inputs in task module are not valid.
        /// </summary>
        /// <param name="groupDetail">groupDetail is the values obtained from task module.</param>
        /// <param name="isSplittingValid">split condition valid flag.</param>
        /// <returns>A task that sends validation card in task module.</returns>
        private MessagingExtensionActionResponse ShowValidationForCreateChannel(GroupDetail groupDetail, bool isSplittingValid)
        {
            return new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        Card = GroupActivityCard.GetGroupActivityValidationCard(groupDetail, isSplittingValid),
                        Height = TaskModuleHeight,
                        Width = TaskModuleWidth,
                    },
                },
            };
        }

        /// <summary>
        /// Send Grouping message in channel from where group activity is invoked.
        /// </summary>
        /// <param name="teamId">Team id of team where bot is installed.</param>
        /// <param name="valuesFromTaskModule">Values obtained from task modules which entered by user.</param>
        /// <param name="membersGroupingWithChannel">A dictionary with members grouped into channels based on entered grouping criteria.</param>
        /// <param name="groupActivityCreator">Team owner who initiated the group activity.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that sends the grouping message in channel.</returns>
        private async Task<string> SendAndStoreGroupingMessageAsync(string groupActivityId, string teamId, GroupDetail valuesFromTaskModule, Dictionary<int, IList<TeamsChannelAccount>> membersGroupingWithChannel, TeamsChannelAccount groupActivityCreator, ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                // Post grouping of members with channel details to channel from where bot is invoked.
                var groupingMessage = this.groupingHelper.GroupingMessage(valuesFromTaskModule, membersGroupingWithChannel, groupActivityCreator.Name);
                var groupingCardActivity = MessageFactory.Attachment(GroupActivityCard.GetGroupActivityCard(groupingMessage, groupActivityCreator.Name, valuesFromTaskModule));
                await turnContext.SendActivityAsync(groupingCardActivity, cancellationToken);

                await this.channelHelper.StoreGroupActivityDetailsAsync(turnContext.Activity.ServiceUrl, groupActivityId, teamId, valuesFromTaskModule, groupActivityCreator.Name, groupingCardActivity.Conversation.Id, groupingCardActivity.Id);
                return groupingMessage;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while creating grouping message for group activity for teamId - {teamId}");
                return null;
            }
        }

        /// <summary>
        /// Method to validate input fields of group activity entered by user.
        /// </summary>
        /// <param name="valuesFromTaskModule">Values obtained from task modules which entered by user.</param>
        /// <returns>Returns true if values entered by user in input fields is valid.</returns>
        private bool IsValidGroupActivityInputFields(GroupDetail valuesFromTaskModule)
        {
            return !string.IsNullOrWhiteSpace(valuesFromTaskModule.GroupTitle) && !string.IsNullOrWhiteSpace(valuesFromTaskModule.GroupDescription)
                    && !string.IsNullOrEmpty(valuesFromTaskModule.SplittingCriteria) && valuesFromTaskModule.ChannelOrMemberUnits >= 2
                    && valuesFromTaskModule.ChannelOrMemberUnits <= 30 && valuesFromTaskModule.DueDate >= DateTime.UtcNow && Validator.HasNoSpecialCharacters(valuesFromTaskModule.GroupTitle);
        }
    }
}
