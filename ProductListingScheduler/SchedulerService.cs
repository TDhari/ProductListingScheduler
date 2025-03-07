using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ProductListingScheduler.API_Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductListingScheduler
{
    public class SchedulerService : BackgroundService
    {
        private readonly ILogger<SchedulerService> _logger;
        private readonly string logPath = @"C:\Users\ADMIN\OneDrive\Desktop\AutomationTrial\Zoho\ErrorLogs\SchedulerLog.txt";
        private readonly string errorLogPath = @"C:\Users\ADMIN\OneDrive\Desktop\AutomationTrial\Zoho\ErrorLogs\ErrorLog.txt";


        public SchedulerService(ILogger<SchedulerService> logger)
        {
            _logger = logger;

        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                await ListProductsOnBestBuy();
                //try
                //{
                //    string currentTime = DateTime.Now.ToString("HH:mm");

                //    if (currentTime == "12:02")
                //    {
                //        await ListProductsOnBestBuy();
                //    }
                //    else if (currentTime == "10:30")
                //    {
                //        await ListProductsOnEbay();
                //    }

                //    await Task.Delay(TimeSpan.FromMinutes(1), stoppingToken); // Check every minute
                //}
                //catch (Exception ex)
                //{
                //    LogError($"Error: {ex.Message}");
                //}
            }
        }


        private async Task ListProductsOnBestBuy()
        {
            Console.WriteLine("................Listing products on BestBuy Method running....................");
            try
            {
                await Task.Delay(2000); // Simulating API call
                IBesBuy bestBuyRepo = new BestBuyRepo();

                BestBuy bestBuy = new BestBuy(bestBuyRepo);
                //await bestBuy.ListProductsOnBestBuy();

                
                bestBuy.SendEmail();

                Log("BestBuy listing completed.");
            }
            catch (Exception ex)
            {
                LogError($"BestBuy Error: {ex.Message}");
            }
        }

        private async Task ListProductsOnEbay()
        {
            try
            {
                await Task.Delay(2000); // Simulating API call
                Log("eBay listing completed.");
            }
            catch (Exception ex)
            {
                LogError($"eBay Error: {ex.Message}");
            }
        }

        private void Log(string message)
        {
            File.AppendAllText(logPath, $"{DateTime.Now}: {message}{Environment.NewLine}");
            _logger.LogInformation(message);
        }

        private void LogError(string message)
        {
            File.AppendAllText(errorLogPath, $"{DateTime.Now}: {message}{Environment.NewLine}");
            _logger.LogError(message);
        }
    }


}
