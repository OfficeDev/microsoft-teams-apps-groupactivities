// <copyright file="BackgroundService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common
{
    using System;
    using System.Globalization;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// BackgroundService class that inherits IHostedService and implements the methods related to background tasks.
    /// </summary>
    public class BackgroundService : IHostedService
    {
        private readonly BackgroundTaskWrapper taskWrapper;
        private readonly ILogger logger;
        private CancellationTokenSource tokenSource;
        private Task currentTask;
        private CultureInfo cultureInfoProvider = CultureInfo.InvariantCulture;

        /// <summary>
        /// Initializes a new instance of the <see cref="BackgroundService"/> class.
        /// </summary>
        /// <param name="taskWrapper">Wrapper class instance for BackgroundTask.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public BackgroundService(BackgroundTaskWrapper taskWrapper, ILogger<BackgroundService> logger)
        {
            this.taskWrapper = taskWrapper;
            this.logger = logger;
        }

        /// <summary>
        /// Method to start the background task when application starts.
        /// </summary>
        /// <param name="cancellationToken">Signals cancellation to the executing method.</param>
        /// <returns>A task instance.</returns>
        public async Task StartAsync(CancellationToken cancellationToken)
        {
            this.logger.LogInformation($"BackgroundService StartAsync method start at: {DateTime.UtcNow.ToString("O", this.cultureInfoProvider)}");

            // Creating a linked token so that we can trigger cancellation outside of this token's cancellation
            this.tokenSource = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            while (cancellationToken.IsCancellationRequested == false)
            {
                try
                {
                    this.logger.LogInformation($"BackgroundService Dequeue method start at: {DateTime.UtcNow.ToString("O", this.cultureInfoProvider)}");

                    // Dequeuing a task and running it in background until the cancellation is triggered or task is completed
                    this.currentTask = this.taskWrapper.Dequeue(this.tokenSource.Token);
                    await this.currentTask;
                }
                catch (OperationCanceledException exception)
                {
                    // Execution has been canceled.
                    this.logger.LogError(exception, "Error in background service");
                }
            }
        }

        /// <summary>
        /// Triggered when the host is performing a graceful shutdown.
        /// </summary>
        /// <param name="cancellationToken">Signals cancellation to the executing method.</param>
        /// <returns>A task instance.</returns>
        public async Task StopAsync(CancellationToken cancellationToken)
        {
            this.logger.LogInformation($"BackgroundService StopAsync method start at: {DateTime.UtcNow.ToString("O", this.cultureInfoProvider)}");

            // Signal cancellation to the executing method
            this.tokenSource.Cancel();

            // If Stop called without start
            if (this.currentTask == null)
            {
                return;
            }

            // Wait until the task completes or the stop token triggers
            await Task.WhenAny(this.currentTask, Task.Delay(-1, cancellationToken));
        }
    }
}
