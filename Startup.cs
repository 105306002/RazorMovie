using System;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Hangfire;
using Hangfire.SqlServer;
using System.Management.Automation;
using System.Collections.ObjectModel;
using Job = Hangfire.Common.Job;

namespace RazorPagesMovie
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }
        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.Configure<CookiePolicyOptions>(options =>
            {
                // This lambda determines whether user consent for non-essential cookies is needed for a given request.
                options.CheckConsentNeeded = context => true;
                options.MinimumSameSitePolicy = SameSiteMode.None;
            });

            // Add Hangfire services.
            services.AddHangfire(configuration => configuration
                .SetDataCompatibilityLevel(CompatibilityLevel.Version_170)
                .UseSimpleAssemblyNameTypeSerializer()
                .UseRecommendedSerializerSettings()
                .UseSqlServerStorage(/*"Server=LAPTOP-C59H4L06;Database=HangfireTest; Integrated Security=True;"*/Configuration.GetConnectionString("HangfireConnection"), new SqlServerStorageOptions
                {
                    CommandBatchMaxTimeout = TimeSpan.FromMinutes(5),
                    SlidingInvisibilityTimeout = TimeSpan.FromMinutes(5),
                    QueuePollInterval = TimeSpan.Zero,
                    UseRecommendedIsolationLevel = true,
                    UsePageLocksOnDequeue = true,
                    DisableGlobalLocks = true
                }));

            // Add the processing server as IHostedService
            services.AddHangfireServer();

            services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_2);
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IBackgroundJobClient backgroundJobs, IHostingEnvironment env)
        {
            //if (env.IsDevelopment())
            //{
            //    app.UseDeveloperExceptionPage();
            //}
            //else
            //{
            //    app.UseExceptionHandler("/Error");
            //    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
            //    app.UseHsts();
            //}

            app.UseHttpsRedirection();
            app.UseStaticFiles();
            app.UseCookiePolicy();


            // 將console的結果呈現在dashboard
            app.UseHangfireDashboard();
            //backgroundJobs.Enqueue(() => Console.WriteLine("Hello world from Hangfire!"));

            //RecurringJob.AddOrUpdate(() => Console.Write("Easy!"), Cron.Minutely);
            //移除排程工作
            RecurringJob.RemoveIfExists("Console.Write"); 

            var manager = new RecurringJobManager();
            manager.AddOrUpdate("2", Job.FromExpression(() => ExecutePowershell()), Cron.Minutely());


            app.UseMvc(routes =>
            {
                routes.MapRoute(
                    name: "default",
                    template: "{controller=Home}/{action=Index}/{id?}");
            });
        }

        public void ExecutePowershell()
        {

                PowerShell PowerShellInstance5 = PowerShell.Create();
                var cmd05 = "Set-Location -Path C:;$currentPath=Get-Location;$context=Get-ChildItem;$context | Out-File -FilePath $currentPath/CurrentDirFileInfo1.txt;";

                PowerShellInstance5.AddScript(cmd05);
                PSDataCollection<PSObject> outputCollection5 = new PSDataCollection<PSObject>();

                //add this:
                outputCollection5.DataAdded += new EventHandler<DataAddedEventArgs>(Output_DataAdded);
                PowerShellInstance5.InvocationStateChanged += new EventHandler<PSInvocationStateChangedEventArgs>(Powershell_InvocationStateChanged);

                PowerShellInstance5.BeginInvoke<PSObject, PSObject>(null, outputCollection5);
        }

        private static void Output_DataAdded(object sender, DataAddedEventArgs e)
        {
            PSDataCollection<PSObject> myp = (PSDataCollection<PSObject>)sender;
            Collection<PSObject> results = myp.ReadAll();
            foreach (PSObject result in results)
            {
                Console.WriteLine(result.ToString());
            }
        }

        private static void Powershell_InvocationStateChanged(object sender, PSInvocationStateChangedEventArgs e)
        {
            Console.WriteLine("PowerShell object state changed: state: {0}\n", e.InvocationStateInfo.State);
        }
    }

}
