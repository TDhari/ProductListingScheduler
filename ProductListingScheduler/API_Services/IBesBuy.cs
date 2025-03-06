using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductListingScheduler.API_Services
{
    interface IBesBuy
    {
        int ListProductOnBestBuy(string filePath);
        void CreateOffersOnBestBuy(string file);
    }
}
