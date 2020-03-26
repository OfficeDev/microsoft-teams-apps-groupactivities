// <copyright file="NotificationHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GroupBot.Cards;
    using Microsoft.Teams.Apps.GroupBot.Common.Interfaces;
    using Microsoft.Teams.Apps.GroupBot.Common.Models;
    using Microsoft.Teams.Apps.GroupBot.Models;
    using Microsoft.Teams.Apps.GroupBot.Models.Configurations;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// Class handles sending notification to channels.
    /// </summary>
    public class NotificationHelper : INotificationHelper
    {
        /// <summary>
        /// Retry policy with jitter, Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </summary>
        private static AsyncRetryPolicy retryPolicy = Policy.Handle<Exception>()
            .WaitAndRetryAsync(Backoff.DecorrelatedJitterBackoffV2(TimeSpan.FromMilliseconds(1000), 2));

        /// <summary>
        /// Helper for storing channel details to azure table storage for sending notification.
        /// </summary>
        private readonly IGroupNotificationStorageHelper groupNotificationStorageHelper;

        /// <summary>
        /// Helper for storing group activity details into storage.
        /// </summary>
        private readonly IGroupActivityStorageHelper groupActivityStorageHelper;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<NotificationHelper> logger;

        /// <summary>
        /// Bot adapter.
        /// </summary>
        private readonly IBotFrameworkHttpAdapter adapter;

        /// <summary>
        /// Tenant id.
        /// </summary>
        private readonly string tenantId;

        /// <summary>
        /// Microsoft app credentials.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Represents a set of key/value application configuration properties for Group activities bot.
        /// </summary>
        private readonly BotAppSetting options;

        /// <summary>
        /// Initializes a new instance of the <see cref="NotificationHelper"/> class.
        /// </summary>
        /// <param name="optionsAccessor">A set of key/value application configuration properties for Group activities bot.</param>
        /// <param name="groupActivityStorageHelper">Helper method for storing group activity details.</param>
        /// <param name="groupNotificationStorageHelper"> Helper method for storing group notification channel details.</param>
        /// <param name="adapter">Bot adapter.</param>
        /// <param name="microsoftAppCredentials">MicrosoftAppCredentials of bot.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public NotificationHelper(IOptionsMonitor<BotAppSetting> optionsAccessor, IGroupActivityStorageHelper groupActivityStorageHelper, IGroupNotificationStorageHelper groupNotificationStorageHelper, IBotFrameworkHttpAdapter adapter, MicrosoftAppCredentials microsoftAppCredentials, ILogger<NotificationHelper> logger)
        {
            this.options = optionsAccessor.CurrentValue;
            this.groupActivityStorageHelper = groupActivityStorageHelper;
            this.groupNotificationStorageHelper = groupNotificationStorageHelper;
            this.logger = logger;
            this.microsoftAppCredentials = microsoftAppCredentials;
            this.adapter = adapter;
            this.tenantId = this.options.TenantId;
        }

        /// <summary>
        /// Method to fetch channel and send notification in channels.
        /// </summary>
        /// <returns>A task that sends notification.</returns>
        public async Task GetChannelsAndSendNotificationAsync()
        {
            this.logger.LogInformation($"Send notification Timer trigger function executed at: {DateTime.UtcNow}");
            try
            {
                this.logger.LogInformation("Get all active group notifications");
                var activeGroupActivities = await this.groupActivityStorageHelper.GetAllActiveGroupNotificationsAsync();
                foreach (GroupActivityEntity groupActivityEntity in activeGroupActivities)
                {
                    var notificationChannels = await this.groupNotificationStorageHelper.GetNotificationChannelsInfoAsync(groupActivityEntity.GroupActivityId);
                    if (notificationChannels.Count > 0)
                    {
                        await this.SendNotificationsAsync(new NotificationRequest
                        {
                            CreatedBy = groupActivityEntity.CreatedBy,
                            DueDate = groupActivityEntity.DueDate,
                            GroupActivityDescription = groupActivityEntity.GroupActivityDescription,
                            GroupActivityTitle = groupActivityEntity.GroupActivityTitle,
                            GroupNotificationChannels = notificationChannels,
                            ServiceUrl = groupActivityEntity.ServiceUrl,
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while getting channels and sending notification");
            }
        }

        /// <summary>
        /// This method is used to send notifications to all channels of a group activity.
        /// </summary>
        /// <param name="request">notification request object.</param>
        /// <returns>Task.</returns>
        private async Task SendNotificationsAsync(NotificationRequest request)
        {
            if (request != null)
            {
                string serviceUrl = request.ServiceUrl;
                MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);
                foreach (GroupNotification channel in request.GroupNotificationChannels)
                {
                    string teamsChannelId = channel.RowKey;
                    var conversationReference = new ConversationReference()
                    {
                        ChannelId = Constants.Channel,
                        Bot = new ChannelAccount() { Id = this.microsoftAppCredentials.MicrosoftAppId },
                        ServiceUrl = serviceUrl,
                        Conversation = new ConversationAccount() { ConversationType = Constants.ChannelConversationType, IsGroup = true, Id = teamsChannelId, TenantId = this.tenantId },
                    };

                    this.logger.LogInformation($"sending notification to channelId- {teamsChannelId}");
                    var card = NotificationCard.GetNotificationCardAttachment(request, channel.ChannelName);
                    try
                    {
                        await retryPolicy.ExecuteAsync(async () =>
                        {
                            await ((BotFrameworkAdapter)this.adapter).ContinueConversationAsync(
                            this.microsoftAppCredentials.MicrosoftAppId,
                            conversationReference,
                            async (conversationTurnContext, conversationCancellationToken) =>
                            {
                                await conversationTurnContext.SendActivityAsync(MessageFactory.Attachment(card));
                            },
                            CancellationToken.None);
                        });
                    }
                    catch (Exception ex)
                    {
                        this.logger.LogError(ex, "Error while sending notification to channel from background service.");
                    }
                }
            }
        }
    }
}
