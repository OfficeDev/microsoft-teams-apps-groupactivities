// <copyright file="BotConnectionSetting.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Models.Configurations
{
    /// <summary>
    /// Provides app setting related to Azure AD bot connection.
    /// </summary>
    public class BotConnectionSetting
    {
        /// <summary>
        /// Gets or sets Azure ADv2 bot connection name.
        /// </summary>
        public string ConnectionName { get; set; }
    }
}
