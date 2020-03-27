// <copyright file="PublicChannelRequest.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Models
{
    using System;
    using Newtonsoft.Json;

    /// <summary>
    /// Handles data for creating public channel.
    /// </summary>
    public class PublicChannelRequest
    {
        /// <summary>
        /// Gets or sets display name of channel in a team.
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        /// <summary>
        /// Gets or sets description of the channel in a team.
        /// </summary>
        [JsonProperty("description")]
        public string Description { get; set; }
    }
}
