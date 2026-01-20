using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Microsoft.Extensions.Logging;

namespace DesktopMcpServer
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var builder = Host.CreateApplicationBuilder(args);

            //Clear default Logging Providers
            builder.Logging.ClearProviders();

            //Add custom File Logger
            string logFilePath = "Logs/DesktopMcpServer-.log";
            builder.Logging.AddProvider(new FileLoggerProvider(logFilePath));

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