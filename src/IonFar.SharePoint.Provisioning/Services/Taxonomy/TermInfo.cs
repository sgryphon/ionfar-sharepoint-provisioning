using System;

namespace IonFar.SharePoint.Provisioning.Services.Taxonomy
{
    public class TermInfo
    {
        public string Name { get; private set; }
        public Guid TermId { get; set; }

        public TermInfo(string name)
        {
            Name = name;
        }
    }
}
