// <copyright file="TeamOwnerValue.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Models.TeamOwnerDetails
{
    using System;
    using Newtonsoft.Json;

    /// <summary>
    /// Handles team owner values as obtained from Microsoft Graph API.
    /// </summary>
    public class TeamOwnerValue
    {
        /// <summary>
        /// Gets or sets odataContext.
        /// </summary>
        [JsonProperty("@odata.type")]
        public string OdataType { get; set; }

        /// <summary>
        /// Gets or sets team owner id.
        /// </summary>
        [JsonProperty("id")]
        public Guid TeamOwnerId { get; set; }

        /// <summary>
        /// Gets or sets team owner display name.
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }
    }
}