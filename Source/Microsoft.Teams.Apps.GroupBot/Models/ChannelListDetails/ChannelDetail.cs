// <copyright file="ChannelDetail.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Models.ChannelListDetails
{
    using Newtonsoft.Json;

    /// <summary>
    /// Channel value obtained from Microsoft Graph API.
    /// </summary>
    public class ChannelDetail
    {
        /// <summary>
        /// Gets or sets display name of Microsoft Teams channel.
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        /// <summary>
        /// Gets or sets membership type. Value is "private" if private channel and "standard" if public channel.
        /// </summary>
        [JsonProperty("membershipType")]
        public string MembershipType { get; set; }
    }
}
