// <copyright file="Validator.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common
{
    using System.Text.RegularExpressions;

    /// <summary>
    /// Class to provide validation methods.
    /// </summary>
    public static class Validator
    {
        /// <summary>
        /// Method the validate special charaters in input string (except space and -).
        /// </summary>
        /// <param name="input">input string.</param>
        /// <returns>Returns true if no special characters found else false.</returns>
        public static bool HasNoSpecialCharacters(string input)
        {
            var regexItem = new Regex("^[a-zA-Z0-9 -]*$");
            return regexItem.IsMatch(input);
        }
    }
}
