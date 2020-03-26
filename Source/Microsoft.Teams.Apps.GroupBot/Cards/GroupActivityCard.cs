// <copyright file="GroupActivityCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Cards
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.GroupBot.Common;
    using Microsoft.Teams.Apps.GroupBot.Models;
    using Microsoft.Teams.Apps.GroupBot.Resources;

    /// <summary>
    /// Class to create new group activity attachments.
    /// </summary>
    public class GroupActivityCard
    {
        /// <summary>
        /// Maximum length allowed for description field of group activity.
        /// </summary>
        private const int DescriptionMaxLength = 500;

        /// <summary>
        /// Maximum length allowed for group title field of group activity.
        /// </summary>
        private const int ChannelMaxLength = 45;

        /// <summary>
        /// Card to render on task module to create group activity.
        /// </summary>
        /// <returns>create new group activity attachment.</returns>
        public static Attachment GetCreateGroupActivityCard()
        {
            AdaptiveCard groupActivityCard = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = Strings.GroupActivityTitle,
                    },
                    new AdaptiveTextInput
                    {
                        Spacing = AdaptiveSpacing.Small,
                        Id = "groupTitle",
                        MaxLength = ChannelMaxLength,
                    },
                    new AdaptiveTextBlock
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Text = Strings.GroupActivityDescriptionTitle,
                    },
                    new AdaptiveTextInput
                    {
                        Spacing = AdaptiveSpacing.Small,
                        Id = "groupDescription",
                        IsMultiline = true,
                        MaxLength = DescriptionMaxLength,
                    },
                    new AdaptiveContainer
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Spacing = AdaptiveSpacing.Medium,
                                                Text = Strings.DueDateText,
                                            },
                                            new AdaptiveDateInput
                                            {
                                                Spacing = AdaptiveSpacing.None,
                                                Id = "dueDate",
                                            },
                                        },
                                    },
                                    new AdaptiveColumn
                                    {
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Spacing = AdaptiveSpacing.Medium,
                                                Text = Strings.DueTimeText,
                                            },
                                            new AdaptiveTimeInput
                                            {
                                                Spacing = AdaptiveSpacing.None,
                                                Id = "dueTime",
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveContainer
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Width = AdaptiveColumnWidth.Stretch,
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Spacing = AdaptiveSpacing.Medium,
                                                Text = Strings.GroupCriteriaText,
                                            },
                                            new AdaptiveChoiceSetInput
                                            {
                                                Spacing = AdaptiveSpacing.Small,
                                                Id = "splittingCriteria",
                                                Choices = new List<AdaptiveChoice>
                                                {
                                                   new AdaptiveChoice
                                                   {
                                                      Title = Strings.SplitInGroupOfGivenMembers,
                                                      Value = Constants.SplitInGroupOfGivenMembers,
                                                   },
                                                   new AdaptiveChoice
                                                   {
                                                      Title = Strings.SplitInGivenNumberOfGroups,
                                                      Value = Constants.SplitInGivenNumberOfGroups,
                                                   },
                                                },
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
                                                Spacing = AdaptiveSpacing.Medium,
                                                Text = Strings.UnitsText,
                                             },
                                             new AdaptiveNumberInput
                                             {
                                                 Spacing = AdaptiveSpacing.Small,
                                                 Id = "units",
                                             },
                                         },
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Text = Strings.AutoCreateChannelQuestionText,
                    },
                    new AdaptiveContainer
                    {
                        Spacing = AdaptiveSpacing.Small,
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
                                            new AdaptiveChoiceSetInput
                                            {
                                                Spacing = AdaptiveSpacing.Small,
                                                Choices = new List<AdaptiveChoice>()
                                                {
                                                    new AdaptiveChoice
                                                    {
                                                        Title = Strings.YesTitle,
                                                        Value = Constants.AutoCreateChannelYes,
                                                    },
                                                    new AdaptiveChoice
                                                    {
                                                        Title = Strings.NoTitle,
                                                        Value = Constants.AutoCreateChannelNo,
                                                    },
                                                },
                                                Id = "autoCreateChannel",
                                                Value = Constants.AutoCreateChannelYes,
                                                Style = AdaptiveChoiceInputStyle.Expanded,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Text = Strings.ChannelTypeQuestionText,
                    },
                    new AdaptiveContainer
                    {
                        Spacing = AdaptiveSpacing.Small,
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveChoiceSetInput
                                            {
                                                Spacing = AdaptiveSpacing.Small,
                                                Choices = new List<AdaptiveChoice>()
                                                {
                                                    new AdaptiveChoice
                                                    {
                                                        Title = Strings.PrivateChannelTypeText,
                                                        Value = Constants.PrivateChannelType,
                                                    },
                                                    new AdaptiveChoice
                                                    {
                                                        Title = Strings.PublicChannelTypeText,
                                                        Value = Constants.PublicChannelType,
                                                    },
                                                },
                                                Id = "channelType",
                                                Value = Constants.PublicChannelType,
                                                Style = AdaptiveChoiceInputStyle.Expanded,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Text = Strings.NotificationQuestionText,
                    },
                    new AdaptiveTextBlock
                    {
                        Spacing = AdaptiveSpacing.None,
                        Text = $"_{Strings.NotificationInformationText}_",
                    },
                    new AdaptiveContainer
                     {
                        Spacing = AdaptiveSpacing.Small,
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveChoiceSetInput
                                            {
                                                Spacing = AdaptiveSpacing.Small,
                                                Choices = new List<AdaptiveChoice>()
                                                {
                                                    new AdaptiveChoice
                                                    {
                                                        Title = Strings.YesTitle,
                                                        Value = Constants.AutoReminderYes,
                                                    },
                                                    new AdaptiveChoice
                                                    {
                                                        Title = Strings.NoTitle,
                                                        Value = Constants.AutoReminderNo,
                                                    },
                                                },
                                                Id = "autoReminder",
                                                Value = Constants.AutoReminderYes,
                                                Style = AdaptiveChoiceInputStyle.Expanded,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                     },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = Strings.SplitButtonText,
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = groupActivityCard,
            };
        }

        /// <summary>
        /// Card to show validations on create group activity card.
        /// </summary>
        /// <param name="groupDetail">group details data as filled on create channel task module.</param>
        /// <param name="isSplittingValid">splitting criteria valid flag.</param>
        /// <returns> An attachment with validation on create new group activity.</returns>
        public static Attachment GetGroupActivityValidationCard(GroupDetail groupDetail, bool isSplittingValid)
        {
            string groupTitleValidationText = string.IsNullOrWhiteSpace(groupDetail.GroupTitle) || !Validator.HasNoSpecialCharacters(groupDetail.GroupTitle) ? Strings.ChannelTitleValidationText : string.Empty;
            string groupDescriptionValidationText = string.IsNullOrWhiteSpace(groupDetail.GroupDescription) ? Strings.ChannelDesciptionValidationText : string.Empty;
            string splittingCriteriaText = string.IsNullOrEmpty(groupDetail.SplittingCriteria) ? Strings.GroupCriteriaValidationText : string.Empty;

            // Not allowing channels to be created if the number of members in each group is exactly 1. Also, not allowing the number of channels creation to be more than 30.
            string unitText = groupDetail.ChannelOrMemberUnits < 2 || groupDetail.ChannelOrMemberUnits > 30 ? Strings.UnitValidationText : string.Empty;
            string dueDateText = groupDetail.DueDate < DateTimeOffset.UtcNow ? Strings.DueDateValidationText : string.Empty;

            if (!isSplittingValid)
            {
                unitText = Strings.GroupCriteriaFailText;
            }

            AdaptiveCard groupActivityValidationCard = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = Strings.GroupActivityTitle,
                    },
                    new AdaptiveTextInput
                    {
                        Spacing = AdaptiveSpacing.Small,
                        Id = "groupTitle",
                        Value = groupDetail.GroupTitle,
                        MaxLength = ChannelMaxLength,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = $"_{groupTitleValidationText}_",
                        Spacing = AdaptiveSpacing.None,
                        IsVisible = !string.IsNullOrEmpty(groupTitleValidationText),
                        Color = AdaptiveTextColor.Attention,
                    },
                    new AdaptiveTextBlock
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Text = Strings.GroupActivityDescriptionTitle,
                    },
                    new AdaptiveTextInput
                    {
                        Spacing = AdaptiveSpacing.Small,
                        Id = "groupDescription",
                        IsMultiline = true,
                        Value = groupDetail.GroupDescription,
                        MaxLength = DescriptionMaxLength,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = $"_{groupDescriptionValidationText}_",
                        Spacing = AdaptiveSpacing.None,
                        IsVisible = !string.IsNullOrEmpty(groupDescriptionValidationText),
                        Color = AdaptiveTextColor.Attention,
                    },
                    new AdaptiveContainer
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Spacing = AdaptiveSpacing.Medium,
                                                Text = Strings.DueDateText,
                                            },
                                            new AdaptiveDateInput
                                            {
                                                Spacing = AdaptiveSpacing.Small,
                                                Id = "dueDate",
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = $"_{dueDateText}_",
                                                Spacing = AdaptiveSpacing.None,
                                                IsVisible = !string.IsNullOrEmpty(dueDateText),
                                                Color = AdaptiveTextColor.Attention,
                                            },
                                        },
                                    },
                                    new AdaptiveColumn
                                    {
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Spacing = AdaptiveSpacing.Medium,
                                                Text = Strings.DueTimeText,
                                            },
                                            new AdaptiveTimeInput
                                            {
                                                Spacing = AdaptiveSpacing.Small,
                                                Id = "dueTime",
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveContainer
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Width = AdaptiveColumnWidth.Stretch,
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Spacing = AdaptiveSpacing.Medium,
                                                Text = Strings.GroupCriteriaText,
                                            },
                                            new AdaptiveChoiceSetInput
                                            {
                                                Spacing = AdaptiveSpacing.Small,
                                                Id = "splittingCriteria",
                                                Choices = new List<AdaptiveChoice>
                                                {
                                                   new AdaptiveChoice
                                                   {
                                                      Title = Strings.SplitInGroupOfGivenMembers,
                                                      Value = Constants.SplitInGroupOfGivenMembers,
                                                   },
                                                   new AdaptiveChoice
                                                   {
                                                      Title = Strings.SplitInGivenNumberOfGroups,
                                                      Value = Constants.SplitInGivenNumberOfGroups,
                                                   },
                                                },
                                                Value = groupDetail.SplittingCriteria,
                                            },
                                            new AdaptiveTextBlock
                                            {
                                                Text = $"_{splittingCriteriaText}_",
                                                Spacing = AdaptiveSpacing.None,
                                                IsVisible = !string.IsNullOrEmpty(splittingCriteriaText),
                                                Color = AdaptiveTextColor.Attention,
                                                Wrap = true,
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
                                                Spacing = AdaptiveSpacing.Medium,
                                                Text = Strings.UnitsText,
                                             },
                                             new AdaptiveNumberInput
                                             {
                                                 Spacing = AdaptiveSpacing.Small,
                                                 Id = "units",
                                                 Value = groupDetail.ChannelOrMemberUnits,
                                             },
                                             new AdaptiveTextBlock
                                             {
                                                Text = $"_{unitText}_",
                                                Spacing = AdaptiveSpacing.None,
                                                IsVisible = !string.IsNullOrEmpty(unitText),
                                                Color = AdaptiveTextColor.Attention,
                                                Wrap = true,
                                             },
                                         },
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Text = Strings.AutoCreateChannelQuestionText,
                    },
                    new AdaptiveContainer
                    {
                        Spacing = AdaptiveSpacing.Small,
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
                                            new AdaptiveChoiceSetInput
                                            {
                                                Spacing = AdaptiveSpacing.Small,
                                                Choices = new List<AdaptiveChoice>()
                                                {
                                                    new AdaptiveChoice
                                                    {
                                                        Title = Strings.YesTitle,
                                                        Value = Constants.AutoCreateChannelYes,
                                                    },
                                                    new AdaptiveChoice
                                                    {
                                                        Title = Strings.NoTitle,
                                                        Value = Constants.AutoCreateChannelNo,
                                                    },
                                                },
                                                Id = "autoCreateChannel",
                                                Value = Constants.AutoCreateChannelYes,
                                                Style = AdaptiveChoiceInputStyle.Expanded,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Text = Strings.ChannelTypeQuestionText,
                    },
                    new AdaptiveContainer
                    {
                        Spacing = AdaptiveSpacing.Small,
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveChoiceSetInput
                                            {
                                                Spacing = AdaptiveSpacing.Small,
                                                Choices = new List<AdaptiveChoice>()
                                                {
                                                    new AdaptiveChoice
                                                    {
                                                        Title = Strings.PrivateChannelTypeText,
                                                        Value = Constants.PrivateChannelType,
                                                    },
                                                    new AdaptiveChoice
                                                    {
                                                        Title = Strings.PublicChannelTypeText,
                                                        Value = Constants.PublicChannelType,
                                                    },
                                                },
                                                Id = "channelType",
                                                Value = Constants.PublicChannelType,
                                                Style = AdaptiveChoiceInputStyle.Expanded,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Text = Strings.NotificationQuestionText,
                    },
                    new AdaptiveTextBlock
                    {
                        Spacing = AdaptiveSpacing.None,
                        Text = $"_{Strings.NotificationInformationText}_",
                    },
                    new AdaptiveContainer
                     {
                        Spacing = AdaptiveSpacing.Small,
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                    new AdaptiveColumn
                                    {
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveChoiceSetInput
                                            {
                                                Spacing = AdaptiveSpacing.Small,
                                                Choices = new List<AdaptiveChoice>()
                                                {
                                                    new AdaptiveChoice
                                                    {
                                                        Title = Strings.YesTitle,
                                                        Value = Constants.AutoReminderYes,
                                                    },
                                                    new AdaptiveChoice
                                                    {
                                                        Title = Strings.NoTitle,
                                                        Value = Constants.AutoReminderNo,
                                                    },
                                                },
                                                Id = "autoReminder",
                                                Value = Constants.AutoReminderYes,
                                                Style = AdaptiveChoiceInputStyle.Expanded,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                     },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = Strings.SplitButtonText,
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = groupActivityValidationCard,
            };
        }

        /// <summary>
        /// Card to render on task module when user is not a team owner.
        /// </summary>
        /// <returns>An attachment to show message that user is not a team owner.</returns>
        public static Attachment GetTeamOwnerErrorCard()
        {
            AdaptiveCard notTeamOwnerCard = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = Strings.NotTeamOwnerText,
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = notTeamOwnerCard,
            };
        }

        /// <summary>
        /// Show grouping details card after successfully grouping the members.
        /// </summary>
        /// <param name="groupingMessage">grouping details of channels with members.</param>
        /// <param name="groupActivityCreator">Team owner who started the group activity.</param>
        /// <param name="groupDetail">Values obtained from task module.</param>
        /// <returns>Returns a card that show new group activity details.</returns>
        public static Attachment GetGroupActivityCard(string groupingMessage, string groupActivityCreator, GroupDetail groupDetail)
        {
            string dueDateString = "{{DATE(" + groupDetail.DueDate.ToUniversalTime().ToString(Constants.Rfc3339DateTimeFormat) + ", SHORT)}}";
            string channelType = groupDetail.ChannelType;
            bool showConclusionText = groupDetail.AutoCreateChannel == Constants.AutoCreateChannelYes;

            AdaptiveCard groupActivityCard = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = string.Format(Strings.GroupCreatorActivityCardText, groupActivityCreator),
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
                                                Text = dueDateString,
                                                Wrap = true,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        Text = Strings.GroupingMessageText,
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = groupingMessage,
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Text = string.Format(Strings.GroupingCardConclusionText, channelType),
                        Wrap = true,
                        IsVisible = showConclusionText,
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = groupActivityCard,
            };
        }

        /// <summary>
        /// Method sends to create attachment in case of some channel failed to be created.
        /// </summary>
        /// <param name="channelsNotCreatedMessage">Channels that are not created.</param>
        /// <param name="groupActivityTitle">Group activity title entered by user for creating new group activity.</param>
        /// <returns>An attachment with channel failure details.</returns>
        public static Attachment GetChannelCreationFailedCard(string channelsNotCreatedMessage, string groupActivityTitle)
        {
            AdaptiveCard channelCreationFailedCard = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = string.Format(Strings.ChannelCreationFailureText, groupActivityTitle),
                    },
                    new AdaptiveTextBlock
                    {
                       Spacing = AdaptiveSpacing.None,
                       Text = Strings.ChannelCreationFailureSubText,
                    },
                    new AdaptiveTextBlock
                    {
                       Spacing = AdaptiveSpacing.None,
                       Text = channelsNotCreatedMessage,
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = channelCreationFailedCard,
            };
        }

        /// <summary>
        /// Show error message in case of required value obtained is null in case of fetch messaging extension action.
        /// </summary>
        /// <returns>An attachment to show generic error message.</returns>
        public static Attachment GetErrorMessageCard()
        {
            AdaptiveCard errorMessageCard = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = Strings.CustomErrorMessage,
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = errorMessageCard,
            };
        }

        /// <summary>
        /// Show error message in case of required value obtained is null in case of fetch messaging extension action.
        /// </summary>
        /// <returns>An attachment to show generic error message.</returns>
        public static Attachment GetTeamNotFoundErrorCard()
        {
            AdaptiveCard errorMessageCard = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = Strings.NoTeamFoundErrorText,
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = errorMessageCard,
            };
        }
    }
}
