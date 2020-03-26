// <copyright file="INotificationHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common.Interfaces
{
    using System.Threading.Tasks;

    /// <summary>
    /// Handles sending notification of group activity in channel.
    /// </summary>
    public interface INotificationHelper
    {
        /// <summary>
        /// Send notification in channel till the due date entered by the user.
        /// </summary>
        /// <returns>A task that sends notification in channel.</returns>
        Task GetChannelsAndSendNotificationAsync();
    }
}