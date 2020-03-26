// <copyright file="TeamOwnerDetails.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Models.TeamOwnerDetails
{
    using System;
    using System.Collections.Generic;
    using Newtonsoft.Json;

    /// <summary>
    /// Handles team owner details obtained from Microsoft Graph API.
    /// </summary>
    public class TeamOwnerDetails
    {
        /// <summary>
        /// Gets or sets odataContext.
        /// </summary>
        [JsonProperty("@odata.context")]
        public Uri OdataContext { get; set; }

        /// <summary>
        /// Gets or sets team details.
        /// </summary>
        [JsonProperty("value")]
        public List<TeamOwnerValue> TeamOwnerValues { get; set; }
    }
}