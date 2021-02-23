using HtmlAgilityPack;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AmazonSeleniumExample
{
    class Program
    {
        static List<string> ASINCodes = new List<string>();
        static DataTable dataTable = new DataTable();



        static void Main(string[] args)
        {

            Console.WriteLine("Amazon bot begins.");

            //That part reads Amazon product ASIN codes that i want to crawl from my local. You can also get some ASIN codes from Amazon URIs.
            string findMaterialPath = @"C:\Users\erkmenesen.AKADEMETRE\Desktop\out\AmazonExample\AmazonASINCodes.txt";

            ASINCodes.AddRange(File.ReadAllLines(findMaterialPath, Encoding.UTF8));

            Console.WriteLine("{0} ASIN Scanned", dataTable.Rows.Count);
            Thread th = new Thread(new ThreadStart(GetProductsDetail));
            th.Start();
        }


        private static void GetProductsDetail()
        {
            ChromeDriver drv = new ChromeDriver();

            Console.WriteLine("Download begins.");
            int wiewed = 0; ;
            int received = 0;
            var products = new List<WebProduct>();
            int repeatCounter = 0;
            foreach (var code in ASINCodes)
            {
                wiewed++;
                try
                {
                    #region Mapping(Seller Detail)
                    Thread.Sleep(3000);
                    repeatCounter = 0;
                Repeat:
                    drv.Navigate().GoToUrl("https://www.amazon.com.tr/s?k=" + code);
                    if (drv.PageSource.IndexOf("Website Temporarily Unavailable") != -1)
                    {
                        if (repeatCounter < 1)
                        {
                            Thread.Sleep(3000);
                            repeatCounter++;
                            goto Repeat;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    if (drv.PageSource.IndexOf("Server Busy") != -1)
                    {
                        if (repeatCounter < 1)
                        {
                            Thread.Sleep(3000);
                            repeatCounter++;
                            goto Repeat;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    try
                    {
                        HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();

                        document.LoadHtml(drv.PageSource.ToString());

                        if (document != null)
                        {
                            HtmlNodeCollection nodCol = document.DocumentNode.SelectNodes("//div[@data-asin]");

                            if (nodCol == null)
                            {
                                WebProduct product = new WebProduct()
                                {
                                    StoreName = "Amazon",
                                    WinCoins = 0,
                                    InStock = true,
                                    StoreId = 1234,
                                    RequestTime = DateTime.Now,
                                    CargoDetail01 = "0",
                                    Barcode = code,
                                    Category = "",
                                    SubCategory = "",
                                    Brand = "",
                                    Sku = "Couldn't find it."
                                };

                                products.Add(product);
                                received++;

                                GC.Collect();
                            }

                            if (nodCol != null)
                            {
                                for (int i = 1; i < nodCol.Count; i++)
                                {
                                    string asin = nodCol[i].Attributes["data-asin"].Value;
                                    if (!string.IsNullOrEmpty(asin))
                                    {
                                        WebProduct product = new WebProduct()
                                        {
                                            StoreName = "Amazon",
                                            WinCoins = 0,
                                            InStock = true,
                                            StoreId = 1234,
                                            RequestTime = DateTime.Now,
                                            CargoDetail01 = "0",
                                            Barcode = code,
                                            Category = "",
                                            SubCategory = "",
                                            Brand = ""
                                        };

                                        try
                                        {
                                            var skudoc = nodCol[i].SelectSingleNode(".//a[@class='a-link-normal a-text-normal']").InnerText.Trim();
                                            product.Sku = skudoc;

                                            var urldoc = nodCol[i].SelectSingleNode(".//a[@class='a-link-normal a-text-normal']").Attributes["href"].Value;
                                            product.NewUrl = "https://amazon.com.tr" + urldoc;
                                        }
                                        catch (Exception ex)
                                        {
                                        }

                                        products.Add(CheckTurkishCharacters(product));
                                        received++;

                                        GC.Collect();
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                    #endregion

                }
                catch (Exception)
                {
                }
                Console.WriteLine("Toplam Ürün : {0}   Bakılan Ürün : {1}  Alınan Ürün(Satıcılı) : {2}", ASINCodes.Count, wiewed, received);

                GC.Collect();
            }

            Console.WriteLine("Data preparing.");
            ConvertToDatatable(products);
            Console.WriteLine("Mission accomplished +respect.");
            drv.Quit();
            Environment.Exit(0);

        }
        private static WebProduct CheckTurkishCharacters(WebProduct model)
        {
            // 1 Characters in text
            WebProduct newModel = new WebProduct(); ;
            var temp = "";
            foreach (PropertyInfo item in model.GetType().GetProperties())
            {
                try
                {
                    if (item.GetValue(model) is string)
                    {
                        temp = item.GetValue(model).ToString();
                        temp = temp.Replace("þ", "ş");
                        temp = temp.Replace("Þ", "Ş");
                        temp = temp.Replace("ý", "ı");
                        temp = temp.Replace("Ý", "İ");
                        temp = temp.Replace("ð", "ğ");
                        temp = temp.Replace("Ð", "Ğ");
                        temp = temp.Replace("&#39;", "'");

                        // Just 2 characters in text;
                        temp = temp.Replace("Ãœ;", "ç");
                        temp = temp.Replace("Ä±", "ı");
                        temp = temp.Replace("ÄŸ", "ğ");
                        temp = temp.Replace("Ã¶", "ö");
                        temp = temp.Replace("ÅŸ", "ş");
                        temp = temp.Replace("Ã¼", "ü");
                        temp = temp.Replace("Ã‡", "Ç");
                        temp = temp.Replace("Ä°", "İ");
                        temp = temp.Replace("ÄŸ", "Ğ");
                        temp = temp.Replace("Ã–", "Ö");
                        temp = temp.Replace("ÅŸ", "Ş");
                        temp = temp.Replace("Ãœ", "Ü");

                        // 5 Characters in text
                        temp = temp.Replace("&#231;", "ç");
                        temp = temp.Replace("&#305;", "ı");
                        temp = temp.Replace("&#287;", "ğ");
                        temp = temp.Replace("&#246;", "ö");
                        temp = temp.Replace("&#351;", "ş");
                        temp = temp.Replace("&#252;", "ü");
                        temp = temp.Replace("&#199;", "Ç");
                        temp = temp.Replace("&#304;", "İ");
                        temp = temp.Replace("&#208;", "Ğ");
                        temp = temp.Replace("&#214;", "Ö");
                        temp = temp.Replace("&#350;", "Ş");
                        temp = temp.Replace("&#220;", "Ü");
                        temp = temp.Replace("&Ccedil;", "Ç");
                        temp = temp.Replace("&#286;", "Ğ");
                        temp = temp.Replace("&Ouml;", "Ö");
                        temp = temp.Replace("&Uuml;", "Ü");
                        temp = temp.Replace("&ccedil;", "ç");
                        temp = temp.Replace("&ouml;", "ö");
                        temp = temp.Replace("&uuml;", "ü");
                        temp = temp.Replace("&amp;", "&");
                        temp = temp.Replace("&middot;", "-");
                    }
                    else
                    {
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                item.SetValue(model, temp);
            }
            return model;
        }

        public static void ConvertToDatatable(List<WebProduct> list)
        {
            try
            {
                DataTable dt = new DataTable();

                string date = DateTime.Now.ToString("yyyy-MM-dd");

                dt.Columns.Add("StoreName");
                dt.Columns.Add("BudgetCategory");
                dt.Columns.Add("Category");
                dt.Columns.Add("SuperGroup");
                dt.Columns.Add("SubCategory");
                dt.Columns.Add("Brand");
                dt.Columns.Add("SKU");
                dt.Columns.Add("SKUCode");
                dt.Columns.Add("Barcode");
                dt.Columns.Add("UnitCode");
                dt.Columns.Add("Supplier");
                dt.Columns.Add("SupplierMark");
                dt.Columns.Add("Price");
                dt.Columns.Add("Stock");
                dt.Columns.Add("IsStock");
                dt.Columns.Add("CargoDetail");
                dt.Columns.Add("CargoPrice");
                dt.Columns.Add("URL");
                dt.Columns.Add("DateTime");


                foreach (var item in list.Distinct().ToList())
                {
                    var row = dt.NewRow();

                    row["StoreName"] = item.StoreName;
                    row["BudgetCategory"] = string.Empty;
                    row["Category"] = item.Category;
                    row["SuperGroup"] = string.Empty;
                    row["SubCategory"] = item.SubCategory;
                    row["Brand"] = item.Brand;
                    row["SKU"] = item.Sku;
                    row["Barcode"] = item.Barcode;
                    row["UnitCode"] = item.Unit;
                    row["SupplierMark"] = item.SupplierMark;
                    row["Supplier"] = item.Supplier;
                    row["Price"] = Convert.ToDouble(item.Price);
                    row["Stock"] = item.StockCount;
                    row["IsStock"] = item.InStock;
                    row["CargoDetail"] = item.CargoDetail01;
                    row["CargoPrice"] = "0";
                    row["URL"] = item.NewUrl;
                    row["DateTime"] = date;

                    dt.Rows.Add(row);
                }

                try
                {
                    GC.Collect();
                    var pathFile = "AmazonKeywordSearch - " + DateTime.Now.ToString("yyyy-MM-dd_HH-mm") + ".xlsx";
                    JazzExcel.ExportToExcel(pathFile, dt, true, false);

                }
                catch (Exception ex)
                {
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
