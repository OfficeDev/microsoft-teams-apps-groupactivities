// <copyright file="ChannelApiResponse.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Models
{
    using System;
    using Newtonsoft.Json;

    /// <summary>
    /// Class for response of public and private channels obtained from Microsoft graph API.
    /// </summary>
    public class ChannelApiResponse
    {
        /// <summary>
        /// Gets or sets odataContext.
        /// </summary>
        [JsonProperty("@odata.context")]
        public Uri OdataContext { get; set; }

        /// <summary>
        /// Gets or sets channel id.
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets channel display name.
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }
    }
}
