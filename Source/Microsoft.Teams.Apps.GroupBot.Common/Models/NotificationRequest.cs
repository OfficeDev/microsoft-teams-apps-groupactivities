// <copyright file="NotificationRequest.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common.Models
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Teams.Apps.GroupBot.Models;

    /// <summary>
    /// Notification request message.
    /// </summary>
    public class NotificationRequest
    {
        /// <summary>
        /// Gets or sets team owner who initiated group activity.
        /// </summary>
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or sets due date of the group activity.
        /// </summary>
        public DateTimeOffset DueDate { get; set; }

        /// <summary>
        /// Gets or sets bot activity service URL.
        /// </summary>
        public string ServiceUrl { get; set; }

        /// <summary>
        /// Gets or sets title of group activity.
        /// </summary>
        public string GroupActivityTitle { get; set; }

        /// <summary>
        /// Gets or sets description of group activity.
        /// </summary>
        public string GroupActivityDescription { get; set; }

        /// <summary>
        /// Gets or sets the notification channel details.
        /// </summary>
        public IList<GroupNotification> GroupNotificationChannels { get; set; }
    }
}
