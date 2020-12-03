using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;

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
                string execlFilePath, downloadsPath;
                Console.Write("Excel file path:");
                execlFilePath = Console.ReadLine();
                //execlFilePath = "C:\\Users\\MBR\\Desktop\\UTiCS\\SearchResults.csv";
                Console.Write("Download folder path:");
                downloadsPath = Console.ReadLine();

                excel = new Excel(execlFilePath, 1);
                int rows = excel.RowsNumbers();


                //get columnNumber which called URL
                int urlColumnNumber = excel.GetColumnNumber("URL");
                int titleColumnNumber = excel.GetColumnNumber("Item Title");
                int pubicationYearColumnNumber = excel.GetColumnNumber("Publication Year");
                int authorColumnNumber = excel.GetColumnNumber("Authors");
                _books = new Book[rows];

                //store files' names, URLs, ... etc form excel into list.
                for (var i = 0; i < rows; i++)
                {
                    _books[i] = new Book();
                    _books[i].URL = excel.ReadCell(i + 1, urlColumnNumber);
                    _books[i].Title = excel.ReadCell(i + 1, titleColumnNumber) + ".pdf";
                    _books[i].PublicationYear = excel.ReadCell(i + 1, pubicationYearColumnNumber);
                    _books[i].Author = excel.ReadCell(i + 1, authorColumnNumber);
                    _books[i].Downloaded = false;
                }

                Console.WriteLine("Preparing URLs.......");
                PrepareDownloadURLs();
                Console.WriteLine("URLs have been prepared.");

                Console.WriteLine("Downloading....");
                DownloadLinks(downloadsPath);
                Console.WriteLine("\nDownload finished.");

                //Console.WriteLine("Check downloads");
                //ComparerFilesAndNames();
                Console.Write("Type 'Enter' to close.");
                Console.ReadLine();
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
        public static bool DownloadLinks(string downloadPath)
        {
            try
            {
                using (var client = new WebClient())
                {
                    try
                    {
                        for (var i = 0; i < _books.Length; i++)
                        {
                            // Console.WriteLine("{0} =>" + _books[i].URL, i);
                            try
                            {
                                client.DownloadFile(_books[i].URL,
                                downloadPath
                                +i+"_"+_books[i].PublicationYear+ "_" + _books[i].Author+ "_" +
                                _books[i].Title +
                                ".pdf");

                            }
                            catch (Exception)
                            {
                                Console.WriteLine("\b\b\bError in file " + i);
                                Console.WriteLine(_books[i].URL);
                            }

                            Console.Write("\b\b\b" + "" + (int)((double)i / _books.Length * 100) + "%");
                        }
                        Console.WriteLine("\b\b\b100%");
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

        // compare the files downloaded with the files in the 
        static void ComparerFilesAndNames()
        {
            string[] filePaths = Directory.GetFiles(@"C:\Users\MBR\Desktop\UTiCS\");
            Console.WriteLine("Number of files {0}", filePaths.Length);
            for (var i = 0; i < filePaths.Length; i++)
            {
                filePaths[i] = filePaths[i].Replace(@"C:\Users\MBR\Desktop\UTiCS\", "");
            }
            for (int i = 0; i < _books.Length; i++)
            {
                if (!filePaths.Contains(_books[i].Title))
                {
                    Console.WriteLine("{0} {1}\n{2}", i + 2, _books[i].Title, _books[i].URL);
                }
            }
        }
    }
}
