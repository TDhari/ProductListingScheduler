using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

//
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Xml.Serialization;

namespace ProductListingScheduler.API_Services
{
    class BestBuyRepo : IBesBuy
    {
        
        
        string apiToken = "1f39fc34-e265-4a28-aaa6-a3427032105b"; // Replace with your actual API token
        string shopid = "2523";
        string badd = "https://marketplace.bestbuy.ca";
        int importid;

        public int ListProductOnBestBuy(string filePath)
        {
            string apiUrl = "https://marketplace.bestbuy.ca/api/products/imports?shop_id="+shopid;
            var operatorFormat = false;

            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(badd.ToString());
                client.DefaultRequestHeaders.Add("Authorization", apiToken);

                using (var content = new MultipartFormDataContent())
                {
                    // Prepare the file content to be uploaded
                    var fileContent = new ByteArrayContent(System.IO.File.ReadAllBytes(filePath));
                    fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    content.Add(fileContent, "file", System.IO.Path.GetFileName(filePath));

                    // Add the shop_id and operator_format to the content if necessary
                    if (!string.IsNullOrWhiteSpace(shopid))
                    {
                        content.Add(new StringContent(shopid), "shop_id");
                    }
                    content.Add(new StringContent(operatorFormat.ToString()), "operator_format");

                    // Execute the POST request to import products
                    var response =  client.PostAsync("/api/products/imports", content);
                    // Ensure a successful response status code
                    response.Result.EnsureSuccessStatusCode();

                    // Process the response
                    var result =  response.Result.Content.ReadAsStringAsync();

                    
                    Console.WriteLine("Response: " + result);
                    importid=result.Id;

                    // Optionally, retrieve the Location header to check the import status
                    if (response.Result.Headers.TryGetValues("Location", out var locationValues))
                    {
                        var location = locationValues.FirstOrDefault();
                        Console.WriteLine("For Products:" + "Check the import status at: " + location);
                    }
                }


            }

            return importid;
               
        }
    }
}
