using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace RazorPagesMovie.Services
{
    public class SecheduleService : BackgroundService
    {
        public SecheduleService(ILoggerFactory loggerFactory)
        {
            Logger = loggerFactory.CreateLogger<SecheduleService>();
        }

        public ILogger Logger { get; }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            Logger.LogInformation("ServiceA is starting.");

            stoppingToken.Register(() => Logger.LogInformation("ServiceA is stopping."));

            while (!stoppingToken.IsCancellationRequested)
            {
                Logger.LogInformation("ServiceA is doing background work.");

                await Task.Delay(TimeSpan.FromSeconds(5), stoppingToken);
            }

            Logger.LogInformation("ServiceA has stopped.");
        }
    }
}
