using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ProductListingScheduler.CommonClasses
{
    class Common
    {

        //category mapping for bestbuy
        public static readonly Dictionary<string, string> categoryMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
         {
            { "laptop", "Computers/Laptops" },
            { "desktop", "Computers/Desktop Computers" },
            { "monitor", "Computers/Monitors" },
            { "tablet", "Computers/iPad & Tablets - Category Branch/Tablets"},
            { "accessories", "Computers/Other Computer & Laptop Accessories"},

        };


        //web hierarchy codes  for bestbuy
        public static readonly Dictionary<string, string> webHierarchymapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
          {

            { "laptop", "BB_36711" },
            { "desktop", "BB_20217" },
            { "monitor", "BB_474833" },
            { "windowstablet", "BB_31040"},
            { "androidtablet", "BB_20356"},
            { "aio", "BB_26199"},

          };

        //product condition mapping for bestbuy
        public static readonly Dictionary<string, string> productConditionMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "new", "Brand New" },
            { "openbox", "Open Box" },
            { "refurbishedexcellent", "Refurbished Excellent" },
            {"refurbishedgood" ,"Refurbished Good"},
            {"refurbishedfair","Refurbished Fair" },

        };

    }
}
