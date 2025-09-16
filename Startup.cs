using System;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.AspNetCore.Http;
using AgentDesktop.Services;
using System.Text.Json.Serialization;
using Microsoft.AspNetCore.HttpOverrides;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.AspNetCore.Mvc.Infrastructure;

namespace AgentDesktop
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

			services.AddMvc(options => options.EnableEndpointRouting = false);

			services.AddControllersWithViews();

			services.AddMemoryCache();

			// Added below Name for session Use
			services.AddMvc().AddSessionStateTempDataProvider();
			services.AddSession(s => s.IdleTimeout = TimeSpan.FromMinutes(30));
			services.AddHttpContextAccessor();	
			


			//services.AddSingleton<IHttpContextAccessor, HttpContextAccessor>();
			services.TryAddSingleton<IActionContextAccessor, ActionContextAccessor>();

			//services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_1);
			services.AddTransient<AgentInterface, AgentRepository>();

			services.AddControllers().AddJsonOptions(options => options.JsonSerializerOptions.Converters.Add(new JsonStringEnumConverter()));

			//services.AddDbContext<CardDBContext>(conn => conn.UseSqlServer(Configuration.GetConnectionString("ExcelConnection")));

			services.Configure<IISServerOptions>(options =>
			{
				options.AllowSynchronousIO = true;
			});

			services.Configure<CookiePolicyOptions>(options =>
			{
				// Set all cookies to be secure
				options.MinimumSameSitePolicy = Microsoft.AspNetCore.Http.SameSiteMode.None;
				options.Secure = Microsoft.AspNetCore.Http.CookieSecurePolicy.Always;
			});


		}

		// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
		public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
		{
			if (env.IsDevelopment())
			{
				app.UseDeveloperExceptionPage();
			}
			else
			{
				app.UseExceptionHandler("/Home/Error");
				// The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
				app.UseHsts();
			}

			app.UseForwardedHeaders(new ForwardedHeadersOptions
			{
				ForwardedHeaders = ForwardedHeaders.XForwardedFor | ForwardedHeaders.XForwardedProto
			});

			app.UseHttpsRedirection();
			app.UseStaticFiles();

			app.UseCookiePolicy();

			app.UseSession();

			app.UseRouting();

			app.UseAuthorization();

			app.UseEndpoints(endpoints =>
			{
				endpoints.MapControllerRoute(
					name: "default",
					pattern: "{controller=Login}/{action=Index}/{id?}");

        });
		}
	}
}
