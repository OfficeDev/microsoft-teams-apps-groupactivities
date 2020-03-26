// <copyright file="GroupNotificationStorageHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GroupBot.Common.Providers;
    using Microsoft.Teams.Apps.GroupBot.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// This is a storage provider having group notification details.
    /// </summary>
    public class GroupNotificationStorageHelper : StorageInitializationHelper, IGroupNotificationStorageHelper
    {
        /// <summary>
        /// Max number of groups for a batch operation.
        /// </summary>
        private const int GroupsPerBatch = 100;
        private const string GroupNotificationTable = "ChannelNotification";

        /// <summary>
        /// Initializes a new instance of the <see cref="GroupNotificationStorageHelper"/> class.
        /// </summary>
        /// <param name="connectionString">storage connection string.</param>
        public GroupNotificationStorageHelper(string connectionString)
            : base(connectionString, GroupNotificationTable)
        {
        }

        /// <summary>
        /// This method returns channel details for a given group activity Id.
        /// </summary>
        /// <param name="groupActivityId">group activity Id.</param>
        /// <returns>list of Group notifications details.</returns>
        public async Task<IList<GroupNotification>> GetNotificationChannelsInfoAsync(string groupActivityId)
        {
            await this.EnsureInitializedAsync();
            string groupCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, groupActivityId);
            TableQuery<GroupNotification> query = new TableQuery<GroupNotification>().Where(groupCondition);
            TableContinuationToken continuationToken = null;
            var groupNotifications = new List<GroupNotification>();

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
        ///  This method is used to insert or replace group notification details.
        /// </summary>
        /// <param name="groupNotifications">groupNotifications list.</param>
        /// <returns>boolean result.</returns>
        public async Task<bool> UpsertGroupNotificationDetailsBatchAsync(IList<GroupNotification> groupNotifications)
        {
            await this.EnsureInitializedAsync();
            TableBatchOperation tableBatchOperation = new TableBatchOperation();
            int batchCount = (int)Math.Ceiling((double)groupNotifications.Count / GroupsPerBatch);
            for (int batchCountIndex = 0; batchCountIndex < batchCount; batchCountIndex++)
            {
                var groupsBatch = groupNotifications.Skip(batchCountIndex * GroupsPerBatch).Take(GroupsPerBatch);
                foreach (var group in groupsBatch)
                {
                    tableBatchOperation.InsertOrReplace(group);
                }

                if (tableBatchOperation.Count > 0)
                {
                    await this.cloudTable.ExecuteBatchAsync(tableBatchOperation);
                }
            }

            return true;
        }
    }
}
