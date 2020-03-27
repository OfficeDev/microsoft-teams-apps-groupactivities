// <copyright file="PrivateChannelRequest.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Models.PrivateChannel
{
    using System.Collections.Generic;
    using Newtonsoft.Json;

    /// <summary>
    /// Model for data for creating private channel.
    /// </summary>
    public class PrivateChannelRequest
    {
        /// <summary>
        /// Gets or sets OdataType.
        /// </summary>
        [JsonProperty("@odata.type")]
        private readonly string oDataType = "#Microsoft.Teams.Core.channel";

        /// <summary>
        /// Gets or sets channel type i.e public or private.
        /// </summary>
        [JsonProperty("membershipType")]
        public string MembershipType { get; set; }

        /// <summary>
        /// Gets or sets channel name.
        /// </summary>
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        /// <summary>
        /// Gets or sets channel description.
        /// </summary>
        [JsonProperty("description")]
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets channel Members.
        /// </summary>
        [JsonProperty("members")]
        public IList<ChannelMember> Members { get; set; }
    }
}
