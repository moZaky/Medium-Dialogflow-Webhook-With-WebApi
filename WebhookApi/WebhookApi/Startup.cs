using System;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Autofac;
using Autofac.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using System.IO;
using Serilog;

namespace DotNetCoreApiSample
{
	public class Startup
	{
		// we use appsettings.json as configuration file to store key-value pairs, 
		// like database connection string, etc
		public IConfiguration configuration { get; }
        public static string wwwRootFolder = string.Empty;

        // we are using Autofac container here to add services and setup Dependency Injection
        public static IContainer container { get; private set; }

		// Constructor: initialize configuration 
		public Startup(IHostingEnvironment env, IConfiguration configuration)
		{
			var builder = new ConfigurationBuilder().
							SetBasePath(env.ContentRootPath).
							AddJsonFile("appsettings.json", false, true).
							AddEnvironmentVariables();
            // Init Serilog configuration
            Log.Logger = new LoggerConfiguration().ReadFrom.Configuration(configuration).CreateLogger();
            configuration = configuration;
            configuration = builder.Build();
		}
       

        // we added MVC and CORS with a new policy named "AllowAll" to allow visiting from any domains
        public IServiceProvider ConfigureServices(IServiceCollection services)
		{
			services.AddMvc();
			services.AddCors(options =>
			{
				options.AddPolicy("AllowAll",
					p => p.AllowAnyOrigin().
						AllowAnyHeader().
						AllowAnyMethod().
						AllowCredentials()
						);
			});
			var builder = new ContainerBuilder();

			builder.Populate(services);

			container = builder.Build();
			return new AutofacServiceProvider(container);
		}

		// We applied the "AllowAll" CORS policy and MVC service 
		// the container would be disposed when the application stopped
		public void Configure(IApplicationBuilder app,
						IHostingEnvironment env,
						ILoggerFactory loggerFctory,
						IApplicationLifetime applicationLifetime)
		{
			app.UseCors("AllowAll");
			app.UseMvc();
            loggerFctory.AddSerilog();
            string baseDir = env.ContentRootPath;
            if (string.IsNullOrWhiteSpace(env.WebRootPath))
            {
                env.WebRootPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
            }
            Startup.wwwRootFolder = env.WebRootPath;

            AppDomain.CurrentDomain.SetData("DataDirectory", System.IO.Path.Combine(baseDir, "App_Data"));
            applicationLifetime.ApplicationStopped.Register(() => container.Dispose());
		}


	}
}