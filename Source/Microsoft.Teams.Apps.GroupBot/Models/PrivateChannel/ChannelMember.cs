// <copyright file="ChannelMember.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Models.PrivateChannel
{
    using System.Collections.Generic;
    using Newtonsoft.Json;

    /// <summary>
    /// Class for members details to be send to create private channel.
    /// </summary>
    public class ChannelMember
    {
        /// <summary>
        /// Gets or sets OdataType.
        /// </summary>
        [JsonProperty("@odata.type")]
        private readonly string oDataType = "#microsoft.graph.aadUserConversationMember";

        /// <summary>
        /// Gets or sets userOdataBind.
        /// </summary>
        [JsonProperty("user@odata.bind")]
        public string UserOdataBind { get; set; }

        /// <summary>
        /// Gets or sets roles.
        /// </summary>
        [JsonProperty("roles")]
        public IList<string> Roles { get; set; }
    }
}
