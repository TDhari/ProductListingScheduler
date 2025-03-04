using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductListingScheduler.CommonClasses
{
    class ReadZohoFile
    {
        public string MasterSKU { get; set; }
        public string SKU { get; set; }
        public string ProductCondition { get; set; }
        public string Model { get; set; }
        public string Brand { get; set; }
        public string Category { get; set; }
        public string SubCategory { get; set; }
        public string Processor { get; set; }
        public string Memory { get; set; }
        public string StorageSSD { get; set; }
        public string GraphicsGPU { get; set; }
        public string OS { get; set; }
        public string ProcessorGeneration { get; set; }
        public int ProcessorCount { get; set; }
        public string ProcessorSpeed { get; set; }
        public string ProcessorSpeedUnit { get; set; }
        public string StorageHDD { get; set; }
        public string HDDInterface { get; set; }
        public string ScreenResolution { get; set; }
        public string OSPlatform { get; set; }
        public string Manufacture { get; set; }
        public string TouchScreen { get; set; }
        public string KeyboardLayout { get; set; }
        public string BacklitKeyboard { get; set; }
        public string OtherFeaturesList { get; set; }
        public string SuitableFor { get; set; }
        public string BundleDetails { get; set; }
        public string BoxContains { get; set; }
        public string ProductSize { get; set; }
        public string ProductSizeUnit { get; set; }
        public decimal RetailPrice { get; set; }
        public string PortsInputOutput { get; set; }
        public decimal ListPriceEbay { get; set; }
        public decimal ListPriceBestBuy { get; set; }
        public decimal ListPriceWalmartCA { get; set; }
        public decimal ListPriceTSC { get; set; }
        public decimal ListPriceAmazon { get; set; }
        public decimal ListPriceEbayMoreDeal { get; set; }
        public decimal ListPriceForRest { get; set; }
        public int QTYAvailable { get; set; }
        public string TAG { get; set; }
        public decimal Width { get; set; }
        public decimal Height { get; set; }
        public decimal Depth { get; set; }
        public string PhysicalDimensionUnit { get; set; }
        public decimal Weight { get; set; }
        public string WeightUnit { get; set; }
        public string Warranty { get; set; }
        public string UPC { get; set; }
        public string ProductTaxCode { get; set; }
        public string WebHierarchy { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string ProcessorGraphics { get; set; }
        public string MemoryType { get; set; }
        public string Color { get; set; }
        public string MPN { get; set; }
        public string DriveType { get; set; }
        public string InputOutputType { get; set; }
        public string Feature1 { get; set; }
        public string Feature2 { get; set; }
        public string Feature3 { get; set; }
        public string Feature4 { get; set; }
        public string WirelessTechnologies { get; set; }
        public decimal ListPriceWalmartUSA { get; set; }
        public string ShortDescription { get; set; }
        public string TitleBestBuy { get; set; }
        public string TitleEbay { get; set; }
        public string TitleTSC { get; set; }
        public string TitleWalmartCA { get; set; }
        public string TitleWalmartUSA { get; set; }
        public string TitleAmazon { get; set; }
        public string NetContent { get; set; }
        public string Image1 { get; set; }
        public string Image2 { get; set; }
        public string Image3 { get; set; }
        public string Image4 { get; set; }
        public string ScreenResolutionForAmazon { get; set; }
        public string GraphicsCardType { get; set; }
    }
}
