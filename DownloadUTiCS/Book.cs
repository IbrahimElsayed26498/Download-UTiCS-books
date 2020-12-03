using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DownloadUTiCS
{
    class Book
    {
        public string URL { get; set; }
        public string Title { get; set; }
        public string PublicationYear{ get; set; }
        public string Author { get; set; }
        public bool Downloaded { get; set; }
    }
}
