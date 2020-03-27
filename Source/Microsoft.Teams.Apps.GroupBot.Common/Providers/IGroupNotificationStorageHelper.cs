// <copyright file="IGroupNotificationStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GroupBot.Models;

    /// <summary>
    /// group notification details storage provider interface.
    /// </summary>
    public interface IGroupNotificationStorageHelper
    {
        /// <summary>
        /// This method returns channel details for a given group activity Id.
        /// </summary>
        /// <param name="groupActivityId">group activity Id.</param>
        /// <returns>list of Group notifications details.</returns>
        Task<IList<GroupNotification>> GetNotificationChannelsInfoAsync(string groupActivityId);

        /// <summary>
        /// Insert or replace group notification details batch in storage.
        /// </summary>
        /// <param name="groupNotifications">GroupNotification list.</param>
        /// <returns>GroupNotification details.</returns>
        Task<bool> UpsertGroupNotificationDetailsBatchAsync(IList<GroupNotification> groupNotifications);
    }
}