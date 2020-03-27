// <copyright file="NotificationBackgroundService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GroupBot.Common
{
    using System;
    using System.Runtime.InteropServices;
    using System.Threading;
    using System.Threading.Tasks;
    using Cronos;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.GroupBot.Common.Interfaces;
    using Microsoft.Win32.SafeHandles;

    /// <summary>
    /// BackgroundService class that inherits IHostedService and implements the methods related to background tasks for sending notification two times a day.
    /// </summary>
    public class NotificationBackgroundService : IHostedService, IDisposable
    {
        private readonly CronExpression expression;
        private readonly TimeZoneInfo timeZoneInfo;
        private readonly ILogger<NotificationBackgroundService> logger;
        private readonly INotificationHelper notificationHelper;
        private readonly BackgroundTaskWrapper backgroundTaskWrapper;
        private System.Timers.Timer timer;
        private int executionCount = 0;

        // Flag: Has Dispose already been called?
        private bool disposed = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="NotificationBackgroundService"/> class.
        /// BackgroundService class that inherits IHostedService and implements the methods related to sending notification tasks.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="notificationHelper">Helper to send notification in channel.</param>
        /// <param name="backgroundTaskWrapper">Wrapper class instance for BackgroundTask.</param>
        public NotificationBackgroundService(ILogger<NotificationBackgroundService> logger, INotificationHelper notificationHelper, BackgroundTaskWrapper backgroundTaskWrapper)
        {
            this.logger = logger;
            this.expression = CronExpression.Parse("0 10,17 * * *");
            this.timeZoneInfo = TimeZoneInfo.Utc;
            this.notificationHelper = notificationHelper;
            this.backgroundTaskWrapper = backgroundTaskWrapper;
        }

        /// <summary>
        /// Method to start the background task when application starts.
        /// </summary>
        /// <param name="cancellationToken">Signals cancellation to the executing method.</param>
        /// <returns>A task instance.</returns>
        public Task StartAsync(CancellationToken cancellationToken)
        {
            this.logger.LogInformation("Notification Hosted Service is running.");
            this.ScheduleNotification();
            return Task.CompletedTask;
        }

        /// <summary>
        /// Triggered when the host is performing a graceful shutdown.
        /// </summary>
        /// <param name="cancellationToken">Signals cancellation to the executing method.</param>
        /// <returns>A task instance.</returns>
        public Task StopAsync(CancellationToken cancellationToken)
        {
            this.logger.LogInformation("Notification Hosted Service is stopping.");
            return Task.CompletedTask;
        }

        /// <summary>
        /// This code added to correctly implement the disposable pattern.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Protected implementation of Dispose pattern.
        /// </summary>
        /// <param name="disposing">True if already disposed else false.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (this.disposed)
            {
                return;
            }

            if (disposing)
            {
                this.timer.Dispose();
            }

            this.disposed = true;
        }

        /// <summary>
        /// Set the timer and enqueue send notification task if timer matched as per Cron expression.
        /// </summary>
        /// <returns>A task that Enqueue sends notification task.</returns>
        private Task ScheduleNotification()
        {
            var count = Interlocked.Increment(ref this.executionCount);
            this.logger.LogInformation("Notification Hosted Service is working. Count: {Count}", count);

            var next = this.expression.GetNextOccurrence(DateTimeOffset.Now, this.timeZoneInfo);
            if (next.HasValue)
            {
                var delay = next.Value - DateTimeOffset.Now;
                this.timer = new System.Timers.Timer(delay.TotalMilliseconds);
                this.timer.Elapsed += (sender, args) =>
                {
                    this.logger.LogInformation($"Timer matched to send notification at timer value : {this.timer}");
                    this.timer.Stop();  // reset timer
                    this.backgroundTaskWrapper.Enqueue(this.SendNotificationAsync()); // Queue the send notification task.
                    this.ScheduleNotification();    // reschedule next
                };
                this.timer.Start();
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// Method invokes send notification task which gets channel name and send the notification.
        /// </summary>
        /// <returns>A task that sends notification in channel for group activity.</returns>
        private async Task SendNotificationAsync()
        {
            this.logger.LogInformation("Notification task queued for sending notification.");
            await this.notificationHelper.GetChannelsAndSendNotificationAsync(); // Send the notifications
        }
    }
}
