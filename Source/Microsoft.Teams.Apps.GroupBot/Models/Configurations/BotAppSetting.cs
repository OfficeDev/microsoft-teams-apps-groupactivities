// <copyright file="BotAppSetting.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Models.Configurations
{
    /// <summary>
    /// Provides app settings related to Group activities bot.
    /// </summary>
    public class BotAppSetting : BotConnectionSetting
    {
        /// <summary>
        /// Gets or sets application base URI.
        /// </summary>
        public string AppBaseURI { get; set; }

        /// <summary>
        /// Gets or sets tenant id.
        /// </summary>
        public string TenantId { get; set; }
    }
}
