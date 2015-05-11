using System.Collections.Generic;
using Microsoft.SharePoint.Client;

namespace IonFar.SharePoint.Provisioning.Services
{
    public class ListDescriptor
    {
        public string WebUrl { get; set; }
        public string ListUrl { get; set; }
        public string ListTitle { get; set; }
        public ListTemplateType ListTemplate { get; set; }
        public List<string> ContentTypeNames { get; set; }
        public bool UseCustomForms { get; set; }
        public string ContentLinkSiteAssetsFileName { get; set; }
        public bool OnQuickLaunch { get; set; }
        public bool IsHidden { get; set; }

        public ListDescriptor()
        {
            OnQuickLaunch = true;
        }
    }
}
