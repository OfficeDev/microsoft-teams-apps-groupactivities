// <copyright file="ChannelListRequest.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Models.ChannelListDetails
{
    using System;
    using System.Collections.Generic;
    using Newtonsoft.Json;

    /// <summary>
    /// Handles channel details obtained from Microsoft Graph API to get channel list.
    /// </summary>
    public class ChannelListRequest
    {
        /// <summary>
        /// Gets or sets odataContext.
        /// </summary>
        [JsonProperty("@odata.context")]
        public Uri OdataContext { get; set; }

        /// <summary>
        /// Gets or sets odataCount.
        /// </summary>
        [JsonProperty("@odata.count")]
        public long OdataCount { get; set; }

        /// <summary>
        /// Gets or sets channel Values.
        /// </summary>
        [JsonProperty("value")]
        public List<ChannelDetail> ChannelsValue { get; set; }
    }
}
