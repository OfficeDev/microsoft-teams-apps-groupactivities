// <copyright file="GroupDetail.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Models
{
    using System;
    using Newtonsoft.Json;

    /// <summary>
    /// Activity details entered by user in task module.
    /// </summary>
    public class GroupDetail
    {
        /// <summary>
        /// Gets or sets title of group activity.
        /// </summary>
        [JsonProperty("groupTitle")]
        public string GroupTitle { get; set; }

        /// <summary>
        /// Gets or sets description of group activity.
        /// </summary>
        [JsonProperty("groupDescription")]
        public string GroupDescription { get; set; }

        /// <summary>
        /// Gets or sets splitting criteria of group activity.
        /// </summary>
        [JsonProperty("splittingCriteria")]
        public string SplittingCriteria { get; set; }

        /// <summary>
        /// Gets or sets units of group activity.
        /// </summary>
        [JsonProperty("units")]
        public int ChannelOrMemberUnits { get; set; }

        /// <summary>
        /// Gets or sets a value indicating auto creation of channel.
        /// </summary>
        [JsonProperty("autoCreateChannel")]
        public string AutoCreateChannel { get; set; }

        /// <summary>
        /// Gets or sets a value indicating channel type that is public channel or private channel.
        /// </summary>
        [JsonProperty("channelType")]
        public string ChannelType { get; set; }

        /// <summary>
        /// Gets or sets due date of group activity.
        /// </summary>
        [JsonProperty("dueDate")]
        public DateTimeOffset DueDate { get; set; }

        /// <summary>
        /// Gets or sets due date of group activity.
        /// </summary>
        [JsonProperty("dueTime")]
        public DateTime DueTime { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether auto reminders of group activity to be send to user.
        /// </summary>
        [JsonProperty("autoReminder")]
        public string AutoReminders { get; set; }
    }
}
