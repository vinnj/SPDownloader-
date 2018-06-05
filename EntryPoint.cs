using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOnlineListDownloader
{
    class EntryPoint
    {
        static void Main(string[] args)
        {
            SPAttachmentDownloader downloader = new SPAttachmentDownloader();
            downloader.DownloadAttachments();
        }
    }
}
