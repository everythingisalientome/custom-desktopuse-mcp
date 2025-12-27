using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace DesktopMcpServer
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var builder = Host.CreateApplicationBuilder(args);

            //Register Service (Singleton)
            builder.Services.AddSingleton<DesktopAutomationService>();

            // Configure Server Options (Name & Version)
            builder.Services.Configure<McpServerOptions>(options =>
            {
                options.ServerInfo = new Implementation
                {
                    Name = "DesktopAutomationMCP",
                    Version = "1.0.0"
                };
            });

            // 3. Add MCP Server
            builder.Services.AddMcpServer()
                   .WithStdioServerTransport()
                   .WithToolsFromAssembly(typeof(Program).Assembly);
            // ^ This scans your project for [McpServerTool] attributes

            var app = builder.Build();
            await app.RunAsync();
        }
    }
}