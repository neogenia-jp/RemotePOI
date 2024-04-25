using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace XlsManiSvc
{
    public class Startup
    {
        public readonly int GRPC_MAX_RECEIVE_MESSAGE_LENGTH = 20 * 1024 * 1024;  // デフォルトの送受信メッセージサイズ上限: 20MB

        // This method gets called by the runtime. Use this method to add services to the container.
        // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=398940
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddGrpc(options =>
            {
                // https://learn.microsoft.com/en-us/aspnet/core/grpc/security?view=aspnetcore-8.0#message-size-limits
                options.MaxReceiveMessageSize = GRPC_MAX_RECEIVE_MESSAGE_LENGTH;
                options.MaxSendMessageSize = GRPC_MAX_RECEIVE_MESSAGE_LENGTH;
            });
            services.AddMemoryCache(option =>
            {
                option.ExpirationScanFrequency = TimeSpan.FromSeconds(5);
            });
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseRouting();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapGrpcService<RemotePOIService>();

                endpoints.MapGet("/", async context =>
                {
                    await context.Response.WriteAsync("Communication with gRPC endpoints must be made through a gRPC client. To learn how to create a client, visit: https://go.microsoft.com/fwlink/?linkid=2086909");
                });
            });
        }
    }
}
