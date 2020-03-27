// <copyright file="WelcomeCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Cards
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.GroupBot.Resources;

    /// <summary>
    /// Methods handles welcome card for group bot.
    /// </summary>
    public static class WelcomeCard
    {
        /// <summary>
        /// Width of the image in the card.
        /// </summary>
        private const string ImageWidth = "1";

        /// <summary>
        /// Width of the content in the card.
        /// </summary>
        private const string ContentWidth = "3";

        /// <summary>
        /// Get welcome card attachment.
        /// </summary>
        /// <param name="appBaseURI">Application base URL.</param>
        /// <returns>An attachment as welcome card.</returns>
        public static Attachment GetWelcomeCardAttachment(string appBaseURI)
        {
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion("1.0"))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = ImageWidth,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = new Uri(string.Format("{0}/images/GroupBotIcon.png", appBaseURI.Trim('/'))),
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Width = ContentWidth,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = Strings.WelcomeCardTitle,
                                        Wrap = true,
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Size = AdaptiveTextSize.Large,
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Text = Strings.WelcomeCardContent,
                                        Wrap = true,
                                        Spacing = AdaptiveSpacing.None,
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
