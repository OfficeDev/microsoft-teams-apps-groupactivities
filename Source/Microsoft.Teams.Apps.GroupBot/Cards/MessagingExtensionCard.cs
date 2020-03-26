// <copyright file="MessagingExtensionCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Web;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.GroupBot.Common;
    using Microsoft.Teams.Apps.GroupBot.Models;
    using Microsoft.Teams.Apps.GroupBot.Resources;

    /// <summary>
    /// Class having method to return messaging extension attachments.
    /// </summary>
    public class MessagingExtensionCard
    {
        /// <summary>
        /// Method to show group activities in messaging extension.
        /// </summary>
        /// <param name="allActivities">All activities of team obtained from table storage.</param>
        /// <returns>A card with activity details and thumbnail card for activity preview.</returns>
        public static List<MessagingExtensionAttachment> GetGroupActivityCard(IList<GroupActivityEntity> allActivities)
        {
            var messagingExtensionAttachments = new List<MessagingExtensionAttachment>();
            foreach (var activity in allActivities)
            {
                string dueDateString = "{{DATE(" + activity.DueDate.ToString(Constants.Rfc3339DateTimeFormat) + ", SHORT)}}";
                string createdOnString = "{{DATE(" + activity.CreatedOn.ToString(Constants.Rfc3339DateTimeFormat) + ", SHORT)}}";

                AdaptiveCard groupActivityCard = new AdaptiveCard("1.0")
                {
                    Body = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Text = activity.GroupActivityTitle,
                            Weight = AdaptiveTextWeight.Bolder,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = $"{activity.CreatedBy} | {createdOnString}",
                            Spacing = AdaptiveSpacing.None,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = Strings.GroupActivityDescriptionTitle,
                            Weight = AdaptiveTextWeight.Bolder,
                        },
                        new AdaptiveTextBlock
                        {
                            Spacing = AdaptiveSpacing.None,
                            Text = activity.GroupActivityDescription,
                            Wrap = true,
                        },
                        new AdaptiveFactSet
                        {
                            Facts = new List<AdaptiveFact>
                            {
                                new AdaptiveFact
                                {
                                    Title = Strings.DueDateText,
                                    Value = dueDateString,
                                },
                            },
                        },
                    },
                    Actions = new List<AdaptiveAction>
                    {
                        new AdaptiveOpenUrlAction
                        {
                            Title = Strings.GoToOriginalThreadButtonText,
                            Url = new Uri(CreateDeeplinkToThread(activity.ConversationId, activity.ActivityId)),
                        },
                    },
                };
                ThumbnailCard previewCard = new ThumbnailCard
                {
                    Title = $"<strong>{HttpUtility.HtmlEncode(activity.GroupActivityTitle)}</strong>",
                    Subtitle = activity.GroupActivityDescription,
                    Text = activity.CreatedBy,
                };

                messagingExtensionAttachments.Add(new Attachment
                {
                    ContentType = AdaptiveCard.ContentType,
                    Content = groupActivityCard,
                }.ToMessagingExtensionAttachment(previewCard.ToAttachment()));
            }

            return messagingExtensionAttachments;
        }

        /// <summary>
        /// Returns go to original thread URI which will help in opening the original conversation about the group activity.
        /// </summary>
        /// <param name="conversationId">The conversation id of group activity card.</param>
        /// <param name="activityId">The activity id of group activity card.</param>
        /// <returns>original thread URI.</returns>
        private static string CreateDeeplinkToThread(string conversationId, string activityId)
        {
            return $"https://teams.microsoft.com/l/message/{conversationId}/{activityId}";
        }
    }
}