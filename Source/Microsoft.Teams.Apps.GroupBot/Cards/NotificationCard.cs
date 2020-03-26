// <copyright file="NotificationCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Text.RegularExpressions;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.GroupBot.Common;
    using Microsoft.Teams.Apps.GroupBot.Common.Models;
    using Microsoft.Teams.Apps.GroupBot.Models;
    using Microsoft.Teams.Apps.GroupBot.Resources;

    /// <summary>
    /// Notification card to be sent to channel.
    /// </summary>
    public class NotificationCard
    {
        /// <summary>
        /// This method gives the notification card attachment to be sent to channel.
        /// </summary>
        /// <param name="notificationRequest">Notification request instance providing notification details.</param>
        /// <param name="channelName">channel name.</param>
        /// <returns>An adaptive card attachment to be send as notification.</returns>
        public static Attachment GetNotificationCardAttachment(NotificationRequest notificationRequest, string channelName)
        {
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion("1.0"))
            {
                Body = new List<AdaptiveElement>
                {
                   new AdaptiveContainer()
                   {
                       Items = new List<AdaptiveElement>()
                       {
                           new AdaptiveTextBlock()
                           {
                               Text = string.Format(Strings.NotificationText, channelName, notificationRequest.CreatedBy),
                               Wrap = true,
                           },
                           new AdaptiveTextBlock()
                           {
                               Text = notificationRequest.GroupActivityTitle,
                               Size = AdaptiveTextSize.Medium,
                               Weight = AdaptiveTextWeight.Bolder,
                               Spacing = AdaptiveSpacing.Medium,
                           },
                           new AdaptiveTextBlock()
                           {
                               Text = notificationRequest.GroupActivityDescription,
                               Wrap = true,
                               Spacing = AdaptiveSpacing.None,
                           },
                           new AdaptiveFactSet()
                           {
                              Facts = new List<AdaptiveFact>()
                              {
                                  new AdaptiveFact()
                                  {
                                       Title = Strings.DueDateText,
                                       Value = "{{DATE(" + notificationRequest.DueDate.ToString(Constants.Rfc3339DateTimeFormat) + ", SHORT)}}",
                                  },
                              },
                           },
                       },
                   },
                },
            };

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
            return adaptiveCardAttachment;
        }

        /// <summary>
        /// Get notification card for new team members in channel.
        /// </summary>
        /// <param name="groupDetail">Group detail.</param>
        /// <param name="createdBy">Group owner name.</param>
        /// <returns>Adaptive card attachment.</returns>
        internal static Attachment GetNewMemberNotificationCard(GroupDetail groupDetail, string createdBy)
        {
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion("1.0"))
            {
                Body = new List<AdaptiveElement>
                {
                   new AdaptiveContainer()
                   {
                       Items = new List<AdaptiveElement>()
                       {
                           new AdaptiveTextBlock()
                           {
                               Text = string.Format(Strings.GroupCreatorActivityCardText, createdBy),
                               Wrap = true,
                           },
                           new AdaptiveContainer
                           {
                               Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveColumnSet
                                    {
                                        Columns = new List<AdaptiveColumn>
                                        {
                                            new AdaptiveColumn
                                            {
                                                Width = AdaptiveColumnWidth.Auto,
                                                Items = new List<AdaptiveElement>
                                                {
                                                    new AdaptiveTextBlock
                                                    {
                                                        Text = $"**{Strings.NameLabel}**",
                                                        Weight = AdaptiveTextWeight.Bolder,
                                                    },
                                                },
                                            },
                                            new AdaptiveColumn
                                            {
                                                Width = AdaptiveColumnWidth.Stretch,
                                                Items = new List<AdaptiveElement>
                                                {
                                                    new AdaptiveTextBlock
                                                    {
                                                        Text = groupDetail.GroupTitle,
                                                        Wrap = true,
                                                    },
                                                },
                                            },
                                        },
                                    },
                                    new AdaptiveColumnSet
                                    {
                                        Columns = new List<AdaptiveColumn>
                                        {
                                            new AdaptiveColumn
                                            {
                                               Width = AdaptiveColumnWidth.Auto,
                                               Items = new List<AdaptiveElement>
                                                {
                                                    new AdaptiveTextBlock
                                                    {
                                                        Text = $"**{Strings.GroupActivityDescriptionTitle}**",
                                                        Weight = AdaptiveTextWeight.Bolder,
                                                    },
                                                },
                                            },
                                            new AdaptiveColumn
                                            {
                                                Width = AdaptiveColumnWidth.Stretch,
                                                Items = new List<AdaptiveElement>
                                                {
                                                    new AdaptiveTextBlock
                                                    {
                                                        Text = groupDetail.GroupDescription,
                                                        Wrap = true,
                                                    },
                                                },
                                            },
                                        },
                                    },
                                    new AdaptiveColumnSet
                                    {
                                        Columns = new List<AdaptiveColumn>
                                        {
                                            new AdaptiveColumn
                                            {
                                                Width = AdaptiveColumnWidth.Auto,
                                                Items = new List<AdaptiveElement>
                                                {
                                                    new AdaptiveTextBlock
                                                    {
                                                        Text = $"**{Strings.DueDateText}**",
                                                        Weight = AdaptiveTextWeight.Bolder,
                                                    },
                                                },
                                            },
                                            new AdaptiveColumn
                                            {
                                                Width = AdaptiveColumnWidth.Stretch,
                                                Items = new List<AdaptiveElement>
                                                {
                                                    new AdaptiveTextBlock
                                                    {
                                                        Text = "{{DATE(" + groupDetail.DueDate.ToString(Constants.Rfc3339DateTimeFormat) + ", SHORT)}}",
                                                        Wrap = true,
                                                    },
                                                },
                                            },
                                        },
                                    },
                                },
                           },
                       },
                   },
                },
            };

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
            return adaptiveCardAttachment;
        }
    }
}
