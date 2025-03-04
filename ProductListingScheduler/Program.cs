// See https://aka.ms/new-console-template for more information
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ProductListingScheduler;
using Microsoft.Extensions.Logging;

Console.WriteLine("Hello, World!");
Console.ReadLine();


var host = Host.CreateDefaultBuilder(args)
    .ConfigureLogging(logging =>
    {
        logging.ClearProviders();
        logging.AddConsole();
    
    
    
    
    
    
    
    })
    .ConfigureServices((context, services) =>
    {
        services.AddHostedService<SchedulerService>();
    })
    .Build();

await host.RunAsync();
