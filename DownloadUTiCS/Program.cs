using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;

namespace DownloadUTiCS
{
    class Program
    {
        static Book[] _books = null;
        static void Main(string[] args)
        {
            Excel excel = null;
            try
            {
                excel = new Excel("C:\\Users\\MBR\\Desktop\\UTiCS\\SearchResults.csv", 1);
                int rows = excel.RowsNumbers();

                //get columnNumber which called URL
                int urlColumnNumber = excel.GetColumnNumber("URL");
                int titleColumnNumber = excel.GetColumnNumber("Item Title");
                _books = new Book[rows];

                //store files' names, and URLs form excel into list.
                for (var i = 0; i < rows; i++)
                {
                    _books[i] = new Book();
                    _books[i].URL = excel.ReadCell(i + 1, urlColumnNumber);
                    _books[i].Title = excel.ReadCell(i + 1, titleColumnNumber);
                    //Console.WriteLine((i + 1) + "->" + _books[i].Title);
                    //Console.WriteLine(_books[i].URL);
                }
                
                Console.WriteLine("Preparing URLs.......");
                PrepareDownloadURLs();
                Console.WriteLine("URLs has been prepared.");

                Console.WriteLine("Downloading....");
                DownloadLinks();
                Console.WriteLine("Downloaeded successfuly.");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
            finally
            {
                excel?.Close();
            }

        }

        public static void PrepareDownloadURLs()
        {
            string url;
            for (var i = 0; i < _books.Length; i++)
            {
                url = _books[i].URL;
                url = url.Replace("book", "content/pdf");
                url = url.Replace("1007/", "1007%2F");
                url = url.Replace("http", "https");
                url += ".pdf";
                _books[i].URL = url;
            }
        }
        public static bool DownloadLinks()
        {
            try
            {
                using (var client = new WebClient())
                {
                    try
                    {
                        for (var i = 98; i < _books.Length; i++)
                        {
                            if (i == 30 || i == 43 || i == 46 || i == 90 || i == 94)
                            {
                                continue;
                            }
                            // Console.WriteLine("{0} =>" + _books[i].URL, i);
                            try
                            {
                                client.DownloadFile(_books[i].URL,
                                "C:\\Users\\MBR\\Desktop\\UTiCS\\" + _books[i].Title + ".pdf");
                                
                            }
                            catch (Exception)
                            {
                                Console.WriteLine("Error in file " + i);
                            }
                            
                            Console.Write("\b\b\b" + "" + (int)((double)i / _books.Length * 100) + "%");
                        }
                    }
                    catch (Exception e)
                    {
                        while (e != null)
                        {
                            Console.WriteLine(e.Message);
                            Console.WriteLine(e.InnerException);
                            e = e.InnerException;
                        }

                    }
                }
                

                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
                throw;
            }
        }
    }
}
