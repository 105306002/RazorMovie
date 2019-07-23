using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.WindowsServices;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using RazorPagesMovie;
using RazorPagesMovie.Services;
using Topshelf;
using NLog.Web;
using Microsoft.Extensions.Logging;
using System;

namespace SampleApp
{
    #region snippet_Program
    public class Program
    {
        public static void Main(string[] args)
        {
            // NLog: setup the logger first to catch all errors
            var logger = NLog.Web.NLogBuilder.ConfigureNLog("nlog.config").GetCurrentClassLogger();
            try
            {
                //logger.Debug("init main");
                //CreateWebHostBuilder(args).Build().Run();

                var isService = !(Debugger.IsAttached || args.Contains("--console"));

                //if (isService)
                //{
                var pathToExe = Process.GetCurrentProcess().MainModule.FileName;
                var pathToContentRoot = Path.GetDirectoryName(pathToExe);
                Directory.SetCurrentDirectory(pathToContentRoot);
                //}

                var builder = CreateWebHostBuilder(
                    args.Where(arg => arg != "--console").ToArray());

                var host = builder.Build();

                //if (isService)
                //{
                // To run the app without the CustomWebHostService change the
                // next line to host.RunAsService();
                //host.RunAsService();
                //}
                //else
                //{
                host.Run();
                //}
            }
            catch (Exception ex)
            {
                //NLog: catch setup errors
                logger.Error(ex, "Stopped program because of exception");
                throw;
            }
            finally
            {
                // Ensure to flush and stop internal timers/threads before application-exit (Avoid segmentation fault on Linux)
                NLog.LogManager.Shutdown();
            }


           
        }

        public static IWebHostBuilder CreateWebHostBuilder(string[] args) =>
            WebHost.CreateDefaultBuilder(args)
                .ConfigureLogging((hostingContext, logging) =>
                {
                    logging.AddEventLog();
                    logging.ClearProviders();
                    logging.SetMinimumLevel(Microsoft.Extensions.Logging.LogLevel.Trace);
                })
                .ConfigureAppConfiguration((context, config) =>
                {
                    // Configure the app here.
                })
                .UseStartup<Startup>()
                .UseNLog();  // NLog: setup NLog for Dependency injection
    }

    #endregion
}