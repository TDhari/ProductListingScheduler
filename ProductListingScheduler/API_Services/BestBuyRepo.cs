using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

//


using Newtonsoft.Json;

namespace ProductListingScheduler.API_Services
{
    class BestBuyRepo : IBesBuy
    {


        string apiToken = "1f39fc34-e265-4a28-aaa6-a3427032105b"; // Replace with your actual API token
        string shopid = "2523";
        string badd = "https://marketplace.bestbuy.ca";
        int importid;

        //List Products on BestBuy
        public int ListProductOnBestBuy(string filePath)
        {
            //string apiUrl = "https://marketplace.bestbuy.ca/api/products/imports?shop_id="+shopid;
            var operatorFormat = false;

            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(badd.ToString());
                client.DefaultRequestHeaders.Add("Authorization", apiToken);

                using (var content = new MultipartFormDataContent())
                {
                    // Prepare the file content to be uploaded
                    var fileContent = new ByteArrayContent(File.ReadAllBytes(filePath));
                    fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    content.Add(fileContent, "file",Path.GetFileName(filePath));

                    // Add the shop_id and operator_format to the content if necessary
                    if (!string.IsNullOrWhiteSpace(shopid))
                    {
                        content.Add(new StringContent(shopid), "shop_id");
                    }
                    content.Add(new StringContent(operatorFormat.ToString()), "operator_format");

                    // Execute the POST request to import products
                    var response = client.PostAsync("/api/products/imports", content);
                    // Ensure a successful response status code
                    response.Result.EnsureSuccessStatusCode();


                    // Process the response
                    var result = response.Result.Content.ReadAsStringAsync();

                    Console.WriteLine("Response: " + result.ToString());

                    importid = response.Id;

                    // Optionally, retrieve the Location header to check the import status
                    if (response.Result.Headers.TryGetValues("Location", out var locationValues))
                    {
                        var location = locationValues.FirstOrDefault();
                        Console.WriteLine("For Products:" + "Check the import status at: " + location);

                        //call the method to check the import status
                        //CheckImportStatus(location);

                        Console.ReadLine();

                    }


                }
            }

            return importid;

        }

        //Create Offers on BestBuy

        public void CreateOffersOnBestBuy(string file)
        {
            var operatorFormat = false;

            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(badd.ToString());
                client.DefaultRequestHeaders.Add("Authorization", apiToken);

                using (var content = new MultipartFormDataContent())
                {
                    // Prepare the file content to be uploaded
                    var fileContent = new ByteArrayContent(File.ReadAllBytes(file));
                    fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    content.Add(fileContent, "file", Path.GetFileName(file));

                    // Add the shop_id and operator_format to the content if necessary
                    if (!string.IsNullOrWhiteSpace(shopid))
                    {
                        content.Add(new StringContent(shopid), "shop_id");
                    }
                    content.Add(new StringContent(operatorFormat.ToString()), "operator_format");
                    content.Add(new StringContent("NORMAL"), "import_mode");

                    // Execute the POST request to import products
                    var response = client.PostAsync("/api/offers/imports", content).Result;

                    // Log the response status code and content
                    Console.WriteLine("Status Code: " + response.StatusCode);
                    var result = response.Content.ReadAsStringAsync().Result;
                    Console.WriteLine("Response: " + result);

                    // Ensure a successful response status code
                    if (response.IsSuccessStatusCode)
                    {
                        //importid = response.id;

                        // Optionally, retrieve the Location header to check the import status
                        if (response.Headers.TryGetValues("Location", out var locationValues))
                        {
                            var location = locationValues.FirstOrDefault();
                            Console.WriteLine("For Products: Check the import status at: " + location);

                            // Call the method to check the import status
                            // CheckImportStatus(location);

                            Console.ReadLine();
                        }
                    }
                    else
                    {
                        Console.WriteLine("Error: " + response.ReasonPhrase);
                    }
                }
            }
        }
        //Check the import status
        public void CheckImportStatus(string url)
        {
            //string url = "https://marketplace.bestbuy.ca/api/products/imports/1523349?shop_id=2523";working fine

            if (!string.IsNullOrEmpty(url))
            {
                using (var client = new HttpClient())
                {
                    //client.BaseAddress = new Uri(badd.ToString());
                    client.DefaultRequestHeaders.Add("Authorization", apiToken);
                    // Execute the GET request to check the import status
                    var response = client.GetAsync(url);
                    // Ensure a successful response status code
                    response.Result.EnsureSuccessStatusCode();
                    // Process the response
                    var result = response.Result.Content.ReadAsStringAsync();
                    Console.WriteLine("Response: " + result.ToString());
                }

            }
            
        }

    }
}
 