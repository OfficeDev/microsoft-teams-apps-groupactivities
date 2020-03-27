// <copyright file="Program.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot
{
    using Microsoft.AspNetCore;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// This a Program  main class for this Bot.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// This main class for this Bot.
        /// </summary>
        /// <param name="args">String of Arguments.</param>
        public static void Main(string[] args)
        {
            CreateWebHostBuilder(args).Build().Run();
        }

        /// <summary>
        /// This method will hit the Startup Method to set up the complete bot services.
        /// </summary>
        /// <param name="args">String of Arguments.</param>
        /// <returns>A unit of Execution.</returns>
        public static IWebHostBuilder CreateWebHostBuilder(string[] args) =>
            WebHost.CreateDefaultBuilder(args)
                .ConfigureLogging((hostingContext, logging) =>
                {
                    // hostingContext.HostingEnvironment can be used to determine environments as well.
                    var appInsightKey = hostingContext.Configuration["AppInsightsInstrumentationKey"];
                    logging.AddApplicationInsights(appInsightKey);
                })
                .UseStartup<Startup>();
    }
}
