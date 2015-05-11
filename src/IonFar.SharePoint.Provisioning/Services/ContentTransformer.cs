using System;
using System.Configuration;

namespace IonFar.SharePoint.Provisioning.Services
{
    /// <summary>
    /// Inspects html, css and js content and replaces links so that they
    /// point to sharepoint environment urls. In addition, also substitutes
    /// environment variables that are specified in javascript pages.
    /// </summary>
    class ContentTransformer
    {
        private Uri _siteUrl;
        private Uri _subsiteUrl;
        private Uri _apiUrl;

        public ContentTransformer(Uri siteUrl, Uri subsiteUrl, Uri apiUrl)
        {
            _siteUrl = siteUrl;
            _subsiteUrl = subsiteUrl;
            _apiUrl = apiUrl;

            if (!apiUrl.AbsolutePath.EndsWith("/api/"))
            {
                throw new ArgumentException("apiUrl");
            }
        }

        public string Fixup(string path)
        {
            var contents = System.IO.File.ReadAllText(path)
                .Replace("url(SiteAssets", "url({rooturl}/SiteAssets")
                .Replace("url(/SiteAssets", "url({rooturl}/SiteAssets")
                .Replace("\"/SiteAssets", "\"{rooturl}/SiteAssets")
                .Replace("'/SiteAssets", "'{rooturl}/SiteAssets")
                .Replace("\"/_layouts", "\"{rooturl}/_layouts")
                .Replace("'/_layouts", "'{rooturl}/_layouts")
                .Replace("ng-include=\"'/SiteAssets", "ng-include=\"'{rooturl}/SiteAssets")
                .Replace("{weburl}", _subsiteUrl.LocalPath)
                .Replace("{rooturl}", _siteUrl.LocalPath)
                .Replace("{apiurl}", _apiUrl.AbsoluteUri);

            return contents;
        }
    }
}
