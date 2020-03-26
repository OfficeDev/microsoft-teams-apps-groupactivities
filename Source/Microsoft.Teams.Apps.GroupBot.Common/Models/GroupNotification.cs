// <copyright file="GroupNotification.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Models
{
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Stores newly created channels details to Azure storage table for sending notification in channels.
    /// </summary>
    public class GroupNotification : TableEntity
    {
        /// <summary>
        /// Gets or sets Channel Id.
        /// </summary>
        public string ChannelId
        {
            get
            {
                return this.RowKey;
            }

            set
            {
                this.RowKey = value;
            }
        }

        /// <summary>
        /// Gets or sets TeamId for which group activity is created.
        /// </summary>
        public string TeamId { get; set; }

        /// <summary>
        /// Gets or sets Channel Name of newly created channel.
        /// </summary>
        public string ChannelName { get; set; }

        /// <summary>
        /// Gets or sets PartionKey.
        /// </summary>
        public string GroupActivityId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }
    }
}
