using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace ProductListingScheduler
{
    class BestBuy
    {
        public async Task ListProductsOnBestBuy()
        {

            //ReadExcel file stored in a folder and get the product details
            List<ReadZohoFile> products = ReadExcelFile();


            foreach (var product in products)
            {
                //write a code to list the product on BestBuy
                Console.WriteLine($"Product listed on BestBuy: {product.SKU}");
                Console.ReadLine();
            }

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

                foreach (var item in columnMapping)
                {
                    Console.WriteLine($"{item.Key} - {item.Value}");
                }
               

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

    
    }
}
