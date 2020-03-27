// <copyright file="IGroupActivityStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GroupBot.Models;

    /// <summary>
    /// Provides the helper methods to store and fetch group activity details into azure table storage.
    /// </summary>
    public interface IGroupActivityStorageHelper
    {
        /// <summary>
        /// Method inserts the newly created group activity details into azure table storage.
        /// </summary>
        /// <param name="groupActivityEntity">Holds newly created group activity details.</param>
        /// <returns>A task that returns true if group activity successfully stored in storage.</returns>
        Task<bool> UpsertGroupActivityAsync(GroupActivityEntity groupActivityEntity);

        /// <summary>
        /// Method to fetch group activity details from azure table storage.
        /// </summary>
        /// <param name="teamId">Team id for which the group activities needs to be fetched.</param>
        /// <returns>A task that returns list of group activities of specified team id.</returns>
        Task<IList<GroupActivityEntity>> GetGroupActivityEntityDetailAsync(string teamId);

        /// <summary>
        /// This method returns all the groups whose notifications are active and due date is not past.
        /// </summary>
        /// <returns>list of Group activity details.</returns>
        Task<IList<GroupActivityEntity>> GetAllActiveGroupNotificationsAsync();
    }
}