#region License, Terms and Author(s)
//
// Lynclog, raw logging for Lync and Skype for business conversations
// Copyright (c) 2016 Philippe Raemy. All rights reserved.
//
//  Author(s):
//
//      Philippe Raemy
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//    http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
#endregion

using System;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace LyncLog
{
    class Program
    {
        private const string AppSettingsFilename = "appsettings.json";
        private const string EnvPrefix = "LYNCLOG_";

        public static async Task Main(string[] args)
        {
            var isService = !(Debugger.IsAttached || args.Contains("--console"));

            var hostBuilder = new HostBuilder()
                .ConfigureHostConfiguration(configBuilder =>
                {
                    configBuilder
                        .AddJsonFile(AppSettingsFilename, optional: true)
                        .AddEnvironmentVariables(prefix: EnvPrefix)
                        .AddCommandLine(args);
                })
                .ConfigureAppConfiguration((hostBuilderCtx, configBuilder) =>
                {
                    configBuilder
                        .AddJsonFile(AppSettingsFilename, optional: true)
                        .AddJsonFile(
                            $"appsettings.{hostBuilderCtx.HostingEnvironment.EnvironmentName}.json",
                            optional: true)
                        .AddEnvironmentVariables(prefix: EnvPrefix)
                        .AddCommandLine(args);
                })
                .ConfigureServices((hostBuilderCtx, services) =>
                {
                    services
                        .AddLogging()
                        .Configure<AppConfig>(hostBuilderCtx.Configuration.GetSection("LyncLog"))
                        // .AddSingleton<IHostedService, LyncLogService>();
                        .AddHostedService<LyncLogService>();
                })
                .ConfigureLogging((hostingContext, logging) =>
                {
                    logging.AddConfiguration(
                        hostingContext.Configuration.GetSection("Logging"));
                    logging.AddConsole();
                })
                .UseConsoleLifetime();

            if (isService)
                await hostBuilder.RunAsServiceAsync();
            else
                await hostBuilder.RunConsoleAsync();
        }
    }
}
