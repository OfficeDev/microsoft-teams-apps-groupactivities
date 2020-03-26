// <copyright file="GroupActivityEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Models
{
    using System;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Stores group activity details to azure table storage.
    /// </summary>
    public class GroupActivityEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets guid id to uniquely identifies the group activity.
        /// </summary>
        public string GroupActivityId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets title of group activity.
        /// </summary>
        public string GroupActivityTitle { get; set; }

        /// <summary>
        /// Gets or sets description of group activity.
        /// </summary>
        public string GroupActivityDescription { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether channel is private or public.
        /// </summary>
        public bool IsPrivateChannel { get; set; }

        /// <summary>
        /// Gets or sets team owner who initiated group activity.
        /// </summary>
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or sets due date of the group activity.
        /// </summary>
        public DateTimeOffset DueDate { get; set; }

        /// <summary>
        /// Gets or sets group activity created date.
        /// </summary>
        public DateTime CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets conversation id of the group activity card that is sent after grouping channel members.
        /// </summary>
        public string ConversationId { get; set; }

        /// <summary>
        /// Gets or sets activity id of the group activity card that is sent after grouping channel members.
        /// </summary>
        public string ActivityId { get; set; }

        /// <summary>
        /// Gets or sets team id for which activities are to be fetched.
        /// </summary>
        public string TeamId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        ///  Gets or sets a value indicating whether Notification is active on channel.
        /// </summary>
        public bool IsNotificationActive { get; set; }

        /// <summary>
        /// Gets or sets bot activity service URL.
        /// </summary>
        public string ServiceUrl { get; set; }
    }
}
