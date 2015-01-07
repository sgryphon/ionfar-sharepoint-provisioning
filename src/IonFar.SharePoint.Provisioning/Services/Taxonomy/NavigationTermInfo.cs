using System;

namespace IonFar.SharePoint.Provisioning.Services.Taxonomy
{
    public class NavigationTermInfo
    {
        public string Url { get; private set; }
        public string Name { get; private set; }
        public Guid TermId { get; set; }

        public NavigationTermInfo(string name, string url)
        {
            Name = name;
            Url = url;
        }
    }
}
