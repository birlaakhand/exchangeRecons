using System;

namespace CefSharp.MinimalExample.Wpf
{
    public class DownloadHandler : IDownloadHandler
    {
        public event EventHandler<DownloadItem> OnBeforeDownloadFired;

        public event EventHandler<DownloadItem> OnDownloadUpdatedFired;

        public void OnBeforeDownload(IBrowser browser, DownloadItem downloadItem, IBeforeDownloadCallback callback)
        {
            Console.WriteLine("Bhai please kaam krja");
            var handler = OnBeforeDownloadFired;
            if (handler != null)
            {
                handler(this, downloadItem);
            }

            if (!callback.IsDisposed)
            {
                using (callback)
                {
                    callback.Continue(@"C:\Users\Akhand\Documents\Visual Studio 2015\Projects\ExchangeRecon\ExchangeRecon\AppFiles\Download Data\" + evalPath(downloadItem), showDialog: false);
                    Console.WriteLine("Suggested Download Path : " + downloadItem.SuggestedFileName);
                    Console.WriteLine("Download URL : " + downloadItem.Url);
                }
            }
        }

        public void OnDownloadUpdated(IBrowser browser, DownloadItem downloadItem, IDownloadItemCallback callback)
        {
            Console.WriteLine("Tu hi kar le bc");
            var handler = OnDownloadUpdatedFired;
            if (handler != null)
            {
                handler(this, downloadItem);
                Console.WriteLine("Shayad Download hua hai");
            }
            if (downloadItem.IsComplete)
                Console.WriteLine(downloadItem.FullPath.ToString() + " ho gayi puri download.");
        }

        public string evalPath(DownloadItem item)
        {
            string itemURL = item.Url;
            if (itemURL.IndexOf("criteria.companyId=4547") > 0)
                return "ICE_CGML.xls";
            else if (itemURL.IndexOf("criteria.companyId=169") > 0)
                return "ICE_CBNA.xls";
            else
                return item.SuggestedFileName;
        }
    }
}