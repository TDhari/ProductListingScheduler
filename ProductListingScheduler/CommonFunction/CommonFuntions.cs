using ProductListingScheduler.CommonClasses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductListingScheduler.CommonFunction
{
    class CommonFuntions
    {

        //get category
        public static string GetCategory(string input)
        {
            return Common.categoryMapping.TryGetValue(input, out string category) ? category : "Unknown Category";
        }

        //get web hierarchy
        public static string GetWebHierarchy(string input)
        {
            return Common.webHierarchymapping.TryGetValue(input, out string webHierarchy) ? webHierarchy : "Unknown Web Hierarchy";
        }

        //get product condition
        public static string GetProductCondition(string input)
        {
            return Common.productConditionMapping.TryGetValue(input, out string productCondition) ? productCondition : "Unknown Product Condition";
        }

        //Convert from inches to cm
        public static double ConvertInchesToCm(double inches)
        {
            return inches * 2.54;
        }

        //Convert from lb to kg
       public static double ConvertLbToKg(double pounds)
        {
            return pounds * 0.453592;
        }

    }
}
