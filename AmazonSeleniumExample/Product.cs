using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AmazonSeleniumExample
{
    public class Product
    {
        public string StoreName { get; set; }
        public string SkuId { get; set; }
        public int StoreId { get; set; }
        public string DbSegment { get; set; }
        public string DbCategory { get; set; }
        public string DbSubCategory { get; set; }
        public int TryNumber { get; set; }
        public DateTime RequestTime { get; set; }
    }
}
