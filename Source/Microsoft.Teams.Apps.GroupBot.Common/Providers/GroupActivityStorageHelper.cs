// <copyright file="GroupActivityStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GroupBot.Common.Providers;
    using Microsoft.Teams.Apps.GroupBot.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage helper which stores group notification and group activity details in Microsoft Azure Table.
    /// </summary>
    public class GroupActivityStorageHelper : StorageInitializationHelper, IGroupActivityStorageHelper
    {
        private const string GroupActivityTable = "GroupActivityMetaData";

        /// <summary>
        /// Initializes a new instance of the <see cref="GroupActivityStorageHelper"/> class.
        /// </summary>
        /// <param name="connectionString">storage connection string.</param>
        public GroupActivityStorageHelper(string connectionString)
            : base(connectionString, GroupActivityTable)
        {
        }

        /// <summary>
        /// Method stores newly created group activity details into storage.
        /// </summary>
        /// <param name="groupActivityEntity">Holds group activity details.</param>
        /// <returns>A task that return true if group activity data is saved or updated.</returns>
        public async Task<bool> UpsertGroupActivityAsync(GroupActivityEntity groupActivityEntity)
        {
            var result = await this.StoreOrUpdateEntityAsync(groupActivityEntity);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Get all group activity details from storage based on team id.
        /// </summary>
        /// <param name="teamId">Team id of team where bot is installed.</param>
        /// <returns>A task that fetches the data from storage. </returns>
        public async Task<IList<GroupActivityEntity>> GetGroupActivityEntityDetailAsync(string teamId)
        {
            await this.EnsureInitializedAsync();
            string teamIdCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, teamId);
            TableQuery<GroupActivityEntity> query = new TableQuery<GroupActivityEntity>().Where(teamIdCondition);
            TableContinuationToken continuationToken = null;
            var groupActivities = new List<GroupActivityEntity>();

            do
            {
                var queryResult = await this.cloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                groupActivities.AddRange(queryResult?.Results);
                continuationToken = queryResult?.ContinuationToken;
            }
            while (continuationToken != null);

            return groupActivities;
        }

        /// <summary>
        /// This method returns all the groups whose notifications are active and due date is not past.
        /// </summary>
        /// <returns>list of Group activity details.</returns>
        public async Task<IList<GroupActivityEntity>> GetAllActiveGroupNotificationsAsync()
        {
            await this.EnsureInitializedAsync();
            string isNotificationActiveCondition = TableQuery.GenerateFilterConditionForBool("IsNotificationActive", QueryComparisons.Equal, true);
            string dueDateCondition = TableQuery.GenerateFilterConditionForDate("DueDate", QueryComparisons.GreaterThanOrEqual, DateTime.UtcNow.Date);
            string combinedFilter = TableQuery.CombineFilters(isNotificationActiveCondition, TableOperators.And, dueDateCondition);
            var query = new TableQuery<GroupActivityEntity>().Where(combinedFilter);
            TableContinuationToken continuationToken = null;
            var groupNotifications = new List<GroupActivityEntity>();

            do
            {
                var queryResult = await this.cloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                groupNotifications.AddRange(queryResult?.Results);
                continuationToken = queryResult?.ContinuationToken;
            }
            while (continuationToken != null);

            return groupNotifications;
        }

        /// <summary>
        /// Stores newly created group activity details into storage.
        /// </summary>
        /// <param name="groupActivityEntity">Holds group activity details and group card conversation id.</param>
        /// <returns>A task that represents group activity entity is saved or updated.</returns>
        private async Task<TableResult> StoreOrUpdateEntityAsync(GroupActivityEntity groupActivityEntity)
        {
            await this.EnsureInitializedAsync();
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(groupActivityEntity);
            return await this.cloudTable.ExecuteAsync(addOrUpdateOperation);
        }
    }
}
