using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Drawing;
using ProductListingScheduler.CommonClasses;
using ProductListingScheduler.API_Services;
using ProductListingScheduler.CommonFunction;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Net;

namespace ProductListingScheduler
{
    class BestBuy
    {
        private readonly IBesBuy _bestBuyRepo;
        public BestBuy(IBesBuy bestBuyRepo)
        {
            this._bestBuyRepo = bestBuyRepo;
        
        }
        public async Task ListProductsOnBestBuy()
        {
            string folderPath = @"C:\Users\ADMIN\OneDrive\Desktop\AutomationTrial\BestBuy";
            string filePath = Path.Combine(folderPath, "BestBuyListingTemplate.xlsx");
            string filePathOffer = Path.Combine(folderPath, "BestBuyOffersTemplate.xlsx");

            //1.ReadExcel file stored in a folder and get the product details
            List<ReadZohoFile> products = ReadExcelFile();

            //2.Send readed data to marketplace template
            WriteToBestBuy(products);


            //Api to send file on marketplace
            int result=_bestBuyRepo.ListProductOnBestBuy(filePath);

            //APi to send offer file on marketplace
            _bestBuyRepo.CreateOffersOnBestBuy(filePathOffer);

            //store file in backup folder
            StoreFileInBackupFolder(filePath);
            StoreFileInBackupFolder(filePathOffer);

            //3.Clean the template file
            CleanExcelFile(filePath);
            //clena offer file
            CleanExcelFile(filePathOffer);

        }

        //Read from excel file logic

        public List<ReadZohoFile> ReadExcelFile()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<ReadZohoFile> products = new List<ReadZohoFile>();

            //write a code to read the excel file located in a folder
            string folderPath = @"C:\Users\ADMIN\OneDrive\Desktop\AutomationTrial\Zoho";

            string filePath = Path.Combine(folderPath, "All Channelsale Feed Exports.xlsx");

            // Ensure the file exists
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("The specified Excel file was not found.", filePath);
            }


            //Read the excel file and get the product details


            // Read the excel file and get the product details
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming the data is in the first worksheet

                // Get the header row
               // var worksheet = package.Workbook.Worksheets[0]; // Read the first worksheet
                int rowCount = worksheet.Dimension.Rows; // Total number of rows
                int colCount = worksheet.Dimension.Columns; // Total number of columns

                // Read header row to map column indexes dynamically
                Dictionary<string, int> columnMapping = new Dictionary<string, int>();
                for (int col = 1; col <= colCount; col++)
                
                {
                    string columnName = worksheet.Cells[1, col].Text.Trim(); // Read header
                    if (!string.IsNullOrEmpty(columnName) && !columnMapping.ContainsKey(columnName))
                    {
                        columnMapping[columnName] = col;
                    }
                }

               // foreach (var item in columnMapping)
              //  {
               //     Console.WriteLine($"{item.Key} - {item.Value}");
              //  }
               

                //Remove

                // Read the rows and columns
                for (int row = 2; row <= rowCount; row++) // Assuming the first row is the header
                {
                    var product = new ReadZohoFile
                    {
                        MasterSKU = GetValue(worksheet, row, columnMapping, "Master SKU"),
                        SKU = GetValue(worksheet, row, columnMapping, "SKU"),
                        ProductCondition = GetValue(worksheet, row, columnMapping, "Product Condition"),
                        Model = GetValue(worksheet, row, columnMapping, "Model"),
                        Brand = GetValue(worksheet, row, columnMapping, "Brand"),
                        Category = GetValue(worksheet, row, columnMapping, "Category"),
                        SubCategory = GetValue(worksheet, row, columnMapping, "Sub Category"),
                        Processor = GetValue(worksheet, row, columnMapping, "Processor"),
                        Memory = GetValue(worksheet, row, columnMapping, "Memory"),
                        StorageSSD = GetValue(worksheet, row, columnMapping, "Storage SSD"),
                        GraphicsGPU = GetValue(worksheet, row, columnMapping, "Graphics/GPU"),
                        OS = GetValue(worksheet, row, columnMapping, "OS"),
                        ProcessorGeneration = GetValue(worksheet, row, columnMapping, "Processor Generation"),
                        ProcessorCount = GetIntValue(worksheet, row, columnMapping, "Processor Count"),
                        ProcessorSpeed = GetValue(worksheet, row, columnMapping, "Processor Speed"),
                        ProcessorSpeedUnit = GetValue(worksheet, row, columnMapping, "Processor Speed Unit"),
                        StorageHDD = GetValue(worksheet, row, columnMapping, "Storage HDD"),
                        HDDInterface = GetValue(worksheet, row, columnMapping, "HDD Interface"),
                        ScreenResolution = GetValue(worksheet, row, columnMapping, "Screen Resolution"),
                        OSPlatform = GetValue(worksheet, row, columnMapping, "OS Platform"),
                        Manufacture = GetValue(worksheet, row, columnMapping, "Manufacture"),
                        TouchScreen = GetValue(worksheet, row, columnMapping, "Touch Screen"),
                        KeyboardLayout = GetValue(worksheet, row, columnMapping, "Keyboard Layout"),
                        BacklitKeyboard = GetValue(worksheet, row, columnMapping, "Backlit Keyboard"),
                        OtherFeaturesList = GetValue(worksheet, row, columnMapping, "Other Features List"),
                        SuitableFor = GetValue(worksheet, row, columnMapping, "Suitable For"),
                        BundleDetails = GetValue(worksheet, row, columnMapping, "Bundle Details"),
                        BoxContains = GetValue(worksheet, row, columnMapping, "Box Contains"),
                        ProductSize = GetValue(worksheet, row, columnMapping, "Product Size"),
                        ProductSizeUnit = GetValue(worksheet, row, columnMapping, "Product Size Unit"),
                        RetailPrice = GetDecimalValue(worksheet, row, columnMapping, "Retail Price"),
                        PortsInputOutput = GetValue(worksheet, row, columnMapping, "Ports (Input/Output)"),
                        ListPriceEbay = GetDecimalValue(worksheet, row, columnMapping, "List Price Ebay"),
                        ListPriceBestBuy = GetDecimalValue(worksheet, row, columnMapping, "List Price BestBuy"),
                        ListPriceWalmartCA = GetDecimalValue(worksheet, row, columnMapping, "List Price for Walmart CA"),
                        ListPriceTSC = GetDecimalValue(worksheet, row, columnMapping, "List Price for TSC"),
                        ListPriceAmazon = GetDecimalValue(worksheet, row, columnMapping, "List Price Amazon"),
                        ListPriceEbayMoreDeal = GetDecimalValue(worksheet, row, columnMapping, "List Price EbayMoreDeal"),
                        ListPriceForRest = GetDecimalValue(worksheet, row, columnMapping, "List Price For Rest"),
                        QTYAvailable = GetIntValue(worksheet, row, columnMapping, "QTY Available"),
                        TAG = GetValue(worksheet, row, columnMapping, "TAG"),
                        Width = GetDecimalValue(worksheet, row, columnMapping, "Width"),
                        Height = GetDecimalValue(worksheet, row, columnMapping, "Height"),
                        Depth = GetDecimalValue(worksheet, row, columnMapping, "Depth"),
                        PhysicalDimensionUnit = GetValue(worksheet, row, columnMapping, "Physical Dimension Unit"),
                        Weight = GetDecimalValue(worksheet, row, columnMapping, "Weight"),
                        WeightUnit = GetValue(worksheet, row, columnMapping, "Weight Unit"),
                        Warranty = GetValue(worksheet, row, columnMapping, "Warranty"),
                        UPC = GetValue(worksheet, row, columnMapping, "UPC"),
                        ProductTaxCode = GetValue(worksheet, row, columnMapping, "Product Tax Code"),
                        WebHierarchy = GetValue(worksheet, row, columnMapping, "Web Hierarchy"),
                        Title = GetValue(worksheet, row, columnMapping, "Title"),
                        Description = GetValue(worksheet, row, columnMapping, "Description"),
                        ProcessorGraphics = GetValue(worksheet, row, columnMapping, "Processor Graphics"),
                        MemoryType = GetValue(worksheet, row, columnMapping, "Memory Type"),
                        Color = GetValue(worksheet, row, columnMapping, "Color"),
                        MPN = GetValue(worksheet, row, columnMapping, "MPN"),
                        DriveType = GetValue(worksheet, row, columnMapping, "Drive Type"),
                        InputOutputType = GetValue(worksheet, row, columnMapping, "Input/Output type"),
                        Feature1 = GetValue(worksheet, row, columnMapping, "Feature1"),
                        Feature2 = GetValue(worksheet, row, columnMapping, "Feature2"),
                        Feature3 = GetValue(worksheet, row, columnMapping, "Feature3"),
                        Feature4 = GetValue(worksheet, row, columnMapping, "Feature4"),
                        WirelessTechnologies = GetValue(worksheet, row, columnMapping, "Wireless Technologies"),
                        ListPriceWalmartUSA = GetDecimalValue(worksheet, row, columnMapping, "List Price Walmart USA"),
                        ShortDescription = GetValue(worksheet, row, columnMapping, "Short Description"),
                        TitleBestBuy = GetValue(worksheet, row, columnMapping, "Title_BestBuy"),
                        TitleEbay = GetValue(worksheet, row, columnMapping, "Title Ebay"),
                        TitleTSC = GetValue(worksheet, row, columnMapping, "Title TSC"),
                        TitleWalmartCA = GetValue(worksheet, row, columnMapping, "Title Walmart CA"),
                        TitleWalmartUSA = GetValue(worksheet, row, columnMapping, "Title Walmart USA"),
                        TitleAmazon = GetValue(worksheet, row, columnMapping, "Title Amazon"),
                        NetContent = GetValue(worksheet, row, columnMapping, "Net Content"),
                        Image1 = GetValue(worksheet, row, columnMapping, "Image1"),
                        Image2 = GetValue(worksheet, row, columnMapping, "Image2"),
                        Image3 = GetValue(worksheet, row, columnMapping, "Image3"),
                        Image4= GetValue(worksheet, row, columnMapping, "Image4"),
                        ScreenResolutionForAmazon = GetValue(worksheet, row, columnMapping, "Screen Resolution For Amazon"),
                        GraphicsCardType = GetValue(worksheet, row, columnMapping, "Graphics Card Type")
                       
                    };

                    products.Add(product);
                }
                //Add the product details to the list
                return products;

            }
        }

        private static string GetValue(ExcelWorksheet ws, int row, Dictionary<string, int> map, string columnName)
        => map.ContainsKey(columnName) ? ws.Cells[row, map[columnName]].Text.Trim() : "";

        private static int GetIntValue(ExcelWorksheet ws, int row, Dictionary<string, int> map, string columnName)
            => map.ContainsKey(columnName) && int.TryParse(ws.Cells[row, map[columnName]].Text.Trim(), out int value) ? value : 0;

        private static decimal GetDecimalValue(ExcelWorksheet ws, int row, Dictionary<string, int> map, string columnName)
            => map.ContainsKey(columnName) && decimal.TryParse(ws.Cells[row, map[columnName]].Text.Trim(), out decimal value) ? value : 0m;

        private static bool GetBoolValue(ExcelWorksheet ws, int row, Dictionary<string, int> map, string columnName)
            => map.ContainsKey(columnName) && ws.Cells[row, map[columnName]].Text.Trim().ToLower() == "yes";

        //Write to Excel File logic For Bestbuy

        public void WriteToBestBuy(List<ReadZohoFile> products)
        {
            string folderPath = @"C:\Users\ADMIN\OneDrive\Desktop\AutomationTrial\BestBuy";
            string filePath = Path.Combine(folderPath, "BestBuyListingTemplate.xlsx");
            string filePathOffer = Path.Combine(folderPath, "BestBuyOffersTemplate.xlsx");

            // Ensure the folder exists
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            FileInfo fileInfo = new FileInfo(filePath);
            FileInfo fileInfoOffer = new FileInfo(filePathOffer);

            using (var package = new ExcelPackage(fileInfo))
            {
                using (var packageOffer = new ExcelPackage(fileInfoOffer))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming the data is in the first worksheet
                    ExcelWorksheet worksheetOffer = packageOffer.Workbook.Worksheets[0]; // Assuming the data is in the first worksheet

                    int startRow = worksheet.Dimension.End.Row + 1; // Start writing after the last row

                    // Write the data rows
                    for (int i = 0; i < products.Count; i++)
                    {
                        var product = products[i];

                        //get category
                        string category = CommonFuntions.GetCategory(product.Category.ToLower().Trim());

                        //get dimensions
                        double height = CommonFuntions.ConvertInchesToCm(Convert.ToDouble(product.Height));
                        double width = CommonFuntions.ConvertInchesToCm(Convert.ToDouble(product.Width));
                        double depth = CommonFuntions.ConvertInchesToCm(Convert.ToDouble(product.Depth));
                        double weight = CommonFuntions.ConvertLbToKg(Convert.ToDouble(product.Weight));

                        //keyboard language
                        string keyboardlanguage = "";
                        if (product.KeyboardLayout.ToLower().Contains("english"))
                        {
                            keyboardlanguage = "English";
                        }
                        else if (product.KeyboardLayout.ToLower().Contains("french"))
                        {
                            keyboardlanguage = "French";
                        }
                        else if (product.KeyboardLayout.ToLower().Contains("bilingual"))
                        {
                            keyboardlanguage = "Bilingual";
                        }
                        else
                        {
                            keyboardlanguage = " ";
                        }
                        //get product condition
                        string productConditionw = Regex.Replace(product.ProductCondition, @"[()\[\]{}]", "");
                        string productCondition = CommonFuntions.GetProductCondition(productConditionw.ToLower().Trim().Replace(" ", ""));

                        //get warranty
                        int warranty = (Convert.ToInt32(product.Warranty)*365);

                        //get today's date
                        string todayDate = DateTimeOffset.Now.ToString("yyyy-MM-ddTHH:mm:ss.fffK");
                        DateTimeOffset now = DateTimeOffset.Now;
                        DateTimeOffset lastDay = new DateTimeOffset(now.Year, now.Month, DateTime.DaysInMonth(now.Year, now.Month), 0, 0, 0, now.Offset);

                        string lastdDateofMonth = lastDay.ToString("yyyy-MM-ddTHH:mm:ss.fffK");

                        int row = startRow + i;
                        worksheet.Cells[row, 1].Value = category;//category
                        worksheet.Cells[row, 2].Value = product.SKU;//sku
                        worksheet.Cells[row, 3].Value = product.TitleBestBuy;//title
                        worksheet.Cells[row, 4].Value = product.ShortDescription;//short description
                        worksheet.Cells[row, 5].Value = product.Brand;//brand
                        worksheet.Cells[row, 6].Value = product.UPC;//upc
                        worksheet.Cells[row, 7].Value = product.Model;//model number
                        worksheet.Cells[row, 8].Value = product.MPN;//MPN
                        worksheet.Cells[row, 9].Value = product.Description;//long decription
                        worksheet.Cells[row, 10].Value = product.Image1;//image 1
                        worksheet.Cells[row, 11].Value = product.Image2;//image 2
                        worksheet.Cells[row, 12].Value = product.Image3;//image 3
                        worksheet.Cells[row, 13].Value = product.Image4;//image 4
                        worksheet.Cells[row, 20].Value = product.WebHierarchy;//web hierarch
                        worksheet.Cells[row, 26].Value = productCondition;// Product condition
                        worksheet.Cells[row, 27].Value = product.BoxContains;//whats in the box
                        worksheet.Cells[row, 28].Value = product.ProductSize;//Screen size
                        worksheet.Cells[row, 29].Value = product.Processor;//Processor type
                        worksheet.Cells[row, 30].Value = product.Memory.ToString().Replace("GB", "").Trim();//Ram Size
                        if (keyboardlanguage != "")
                        {
                            worksheet.Cells[row, 34].Value = keyboardlanguage;//--Keyboard Language
                        }

                        worksheet.Cells[row, 35].Value = product.BacklitKeyboard;//Backlit Keyboard

                        worksheet.Cells[row, 36].Value = product.PortsInputOutput;//other input/output ports
                                                                                  //worksheet.Cells[row, 37].Value = product.OSPlatform;//product platform
                        worksheet.Cells[row, 43].Value = product.Color;//color
                        worksheet.Cells[row, 47].Value = product.OS;//operating system
                        worksheet.Cells[row, 49].Value = product.ProcessorSpeed;//Processor Spee
                                                                                //worksheet.Cells[row, 54].Value = product.ProcessorGraphics;//Operating System Language
                        worksheet.Cells[row, 56].Value = product.ProcessorCount;//Processor Core
                        worksheet.Cells[row, 57].Value = product.TouchScreen;//Touch screen Display 
                        worksheet.Cells[row, 61].Value = product.StorageHDD.Replace("GB", "").Trim();//Hard disk drive capacity
                        worksheet.Cells[row, 62].Value = product.StorageSSD.Replace("GB", "").Trim();//SSD capacity
                        worksheet.Cells[row, 63].Value = product.GraphicsGPU;//Graphics card
                        worksheet.Cells[row, 66].Value = width;//width
                        worksheet.Cells[row, 67].Value = height;//Height
                        worksheet.Cells[row, 68].Value = depth;//Depth
                        worksheet.Cells[row, 69].Value = weight;//Weight
                        worksheet.Cells[row, 81].Value = product.MemoryType;//RAM Type
                        worksheet.Cells[row, 90].Value = product.ProcessorGeneration;//Processor Generation/Serie
                        worksheet.Cells[row, 92].Value = product.ProductCondition;//Product condition
                        worksheet.Cells[row, 110].Value = product.Processor;//Processor type
                        worksheet.Cells[row, 145].Value = product.ScreenResolution;//Resolution


                        //write to offer file
                        worksheetOffer.Cells[row, 1].Value = product.SKU;//sku
                        worksheetOffer.Cells[row, 2].Value = product.UPC;//id
                        worksheetOffer.Cells[row, 3].Value = "UPC-A";//id type
                        worksheetOffer.Cells[row, 6].Value = product.RetailPrice;//price
                        worksheetOffer.Cells[row, 8].Value = product.QTYAvailable;//quantity
                        worksheetOffer.Cells[row, 9].Value = 2;//condition
                        worksheetOffer.Cells[row, 10].Value = "New";//condition type
                        worksheetOffer.Cells[row, 11].Value = todayDate;//avalabbility start date
                        //worksheetOffer.Cells[row, 12].Value = "2099-12-31";//avalabbility end date
                        worksheetOffer.Cells[row, 14].Value = product.ListPriceBestBuy;//discount price
                        worksheetOffer.Cells[row, 15].Value = todayDate;//discount start date
                        worksheetOffer.Cells[row, 16].Value = lastdDateofMonth;//discount end date
                        worksheetOffer.Cells[row, 19].Value = warranty;//warranty in days



                    }

                    // Save the Excel file
                    package.Save();
                    packageOffer.Save();                }

            }
        }


        //Clean the bestbuy template file
        public void CleanExcelFile(string filePath)
        {
            FileInfo fileInfo = new FileInfo(filePath);

            using (var package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming the data is in the first worksheet

                int startRow = 3; // Start cleaning from the 4th row
                int endRow = worksheet.Dimension.End.Row; // Get the last row

                
                if (endRow >= 3)  // Ensure there are at least 3 rows
                {
                    worksheet.DeleteRow(3, endRow - 2); // Delete rows starting from 3rd row
                }

                // Save the changes
                package.Save();
            }
        }


        //store file in backup folder

        public void StoreFileInBackupFolder(string filePath)
        {
            string folderPath = @"C:\Users\ADMIN\OneDrive\Desktop\AutomationTrial\BestBuyBackup";

            if (filePath.Contains("Offer"))
            {
                 folderPath = @"C:\Users\ADMIN\OneDrive\Desktop\AutomationTrial\BestBuyBackup\Offers";
            }
            else
            {
                 folderPath = @"C:\Users\ADMIN\OneDrive\Desktop\AutomationTrial\BestBuyBackup\Products";
            }
                
            //string fileName = Path.GetFileName(filePath);
            //string backupFilePath = Path.Combine(folderPath, fileName);
            // Ensure the folder exists
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            // Generate a timestamped backup filename
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string backupFileName = $"{fileName}_Backup_{timestamp}{extension}";
            string backupFilePath = Path.Combine(folderPath, backupFileName);

            // Copy file to backup location
            File.Copy(filePath, backupFilePath, true);

        }

        //send email
        public void SendEmail()
        {

            try
            {
                // Email account configuration
                string smtpServer = "smtp.zoho.com"; // Change based on your provider
                int smtpPort = 587; // Use 465 for SSL, 587 for TLS
                string senderEmail = "twinkle@dhari.ca";
                string senderPassword = "password@dhari"; // Use app password if required
                string recipientEmail = "twinkle@dhari.ca";

                // Create email message
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(senderEmail);
                mail.To.Add(recipientEmail);
                mail.Subject = "Test Email from twinkle@dhari.ca";
                mail.Body = "Hello, this is a test email sent from my C# application.";
                mail.IsBodyHtml = false; // Set to true if sending HTML content

                // Configure SMTP client
                SmtpClient smtpClient = new SmtpClient(smtpServer, smtpPort)
                {
                    Credentials = new NetworkCredential(senderEmail, senderPassword),
                    EnableSsl = true // Set to false if SSL is not required
                };

                // Send email
                smtpClient.Send(mail);
                Console.WriteLine("Email sent successfully from twinkle@dhari.ca.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error sending email: " + ex.Message);
            }
        }


    }
}
