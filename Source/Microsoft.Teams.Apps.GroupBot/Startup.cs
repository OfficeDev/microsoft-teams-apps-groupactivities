// <copyright file="Startup.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot
{
    using System;
    using System.Collections.Generic;
    using System.Net.Http;
    using System.Text;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Authentication.JwtBearer;
    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Azure;
    using Microsoft.Bot.Builder.BotFramework;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.IdentityModel.Tokens;
    using Microsoft.Teams.Apps.GroupBot.Bots;
    using Microsoft.Teams.Apps.GroupBot.Common;
    using Microsoft.Teams.Apps.GroupBot.Common.Interfaces;
    using Microsoft.Teams.Apps.GroupBot.Models.Configurations;
    using Polly;
    using Polly.Extensions.Http;

    /// <summary>
    /// Startup class.
    /// </summary>
    public class Startup
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Startup"/> class.
        /// </summary>
        /// <param name="configuration">Startup Configuration.</param>
        public Startup(IConfiguration configuration)
        {
            this.Configuration = configuration;
        }

        /// <summary>
        /// Gets Configurations Interfaces.
        /// </summary>
        public IConfiguration Configuration { get; }

        /// <summary>
        /// This method gets called by the runtime. Use this method to add services to the container.
        /// </summary>
        /// <param name="services">Service Collection Interface.</param>
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddSingleton<TelemetryClient>();

            services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_1);

            services.AddHttpClient<IGraphApiHelper, GraphApiHelper>().AddPolicyHandler(GetRetryPolicy());

            services.AddHostedService<NotificationBackgroundService>();
            services.AddSingleton<INotificationHelper, NotificationHelper>();

            services.AddHostedService<BackgroundService>();
            services.AddSingleton<BackgroundTaskWrapper>();

            services.AddSingleton<IChannelHelper, ChannelHelper>();

            services.AddSingleton<IGroupingHelper, GroupingHelper>();

            services.AddSingleton<IGroupNotificationStorageHelper>(new GroupNotificationStorageHelper(this.Configuration["StorageConnectionString"]));

            services.AddSingleton<IGroupActivityStorageHelper>(new GroupActivityStorageHelper(this.Configuration["StorageConnectionString"]));

            services.AddSingleton<ITeamUserHelper, TeamUserHelper>();

            services.AddSingleton<ICredentialProvider, ConfigurationCredentialProvider>();

            // Create the Bot Framework Adapter with error handling enabled.
            services.AddSingleton<IBotFrameworkHttpAdapter, AdapterWithErrorHandler>();

            // App insights telemetry.
            services.AddSingleton<TelemetryClient>();

            services.AddSingleton(new OAuthClient(new MicrosoftAppCredentials(this.Configuration["MicrosoftAppId"], this.Configuration["MicrosoftAppPassword"])));

            // Microsoft app credentials
            services.AddSingleton(new MicrosoftAppCredentials(this.Configuration["MicrosoftAppId"], this.Configuration["MicrosoftAppPassword"]));

            services.Configure<BotConnectionSetting>(this.Configuration);

            services.Configure<BotAppSetting>(this.Configuration);

            // Create the bot as a transient. In this case the ASP Controller is expecting an IBot.
            services.AddTransient<IBot, GroupActivitiesBot>();
        }

        /// <summary>
        /// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        /// </summary>
        /// <param name="app">Application Builder.</param>
        /// <param name="env">Hosting Environment.</param>
        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseHsts();
            }

            app.UseAuthentication();
            app.UseDefaultFiles();
            app.UseStaticFiles();
            app.UseMvc();
        }

        /// <summary>
        /// Retry policy for transient error cases.
        /// If there is no success code in response, request will be sent again for two times
        /// with interval of 2 and 8 seconds respectively.
        /// </summary>
        /// <returns>Policy.</returns>
        private static IAsyncPolicy<HttpResponseMessage> GetRetryPolicy()
        {
            return HttpPolicyExtensions
                .HandleTransientHttpError()
                .OrResult(response => response.IsSuccessStatusCode == false)
                .WaitAndRetryAsync(3, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)));
        }
    }
}
