// <copyright file="Constants.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common
{
    /// <summary>
    /// Constants class.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// Sets the private channel type.
        /// </summary>
        public const string PrivateChannelType = "private";

        /// <summary>
        /// Sets the private channel type.
        /// </summary>
        public const string PublicChannelType = "public";

        /// <summary>
        /// Split channel based on criteria of equal group choice.
        /// </summary>
        public const string SplitInGroupOfGivenMembers = "SplitInGroupOf";

        /// <summary>
        /// Split channel based on criteria of equal number of group.
        /// </summary>
        public const string SplitInGivenNumberOfGroups = "SplitInNumberofGroups";

        /// <summary>
        /// Auto create channels after grouping members of the team.
        /// </summary>
        public const string AutoCreateChannelYes = "Yes";

        /// <summary>
        /// Only create groups of the team members.Do not create channels.
        /// </summary>
        public const string AutoCreateChannelNo = "No";

        /// <summary>
        /// Send notification in channel after creating group activity.
        /// </summary>
        public const string AutoReminderYes = "Yes";

        /// <summary>
        /// Send notification in channel after creating group activity.
        /// </summary>
        public const string AutoReminderNo = "No";

        /// <summary>
        /// default value for channel activity to send notifications.
        /// </summary>
        public const string Channel = "msteams";

        /// <summary>
        /// Channel conversation type to send notification.
        /// </summary>
        public const string ChannelConversationType = "channel";

        /// <summary>
        /// Microsoft Graph API base URI.
        /// </summary>
        public const string GraphAPIBaseURL = "https://graph.microsoft.com";

        /// <summary>
        /// Date format for cards.
        /// </summary>
        public const string Rfc3339DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'";
    }
}
