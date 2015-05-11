using System.Collections.Generic;

namespace IonFar.SharePoint.Provisioning.Services
{
    public sealed class ContentTypeDescriptor
    {
        public ContentTypeDescriptor(
            string id,
            string name,
            string description,
            string group,
            ContentTypeFieldReference[] fields
        ) 
        {
            Id = id;
            Name = name;
            Description = description;
            Group = group;
            Fields = fields;
        }

        public string Id { get; private set; }
        public string Name { get; private set; }
        public string Description { get; private set; }
        public string Group { get; private set; }
        public IEnumerable<ContentTypeFieldReference> Fields { get; private set; }
    } 

    public enum ContentTypeFieldStatus
    {
        Required,
        Optional,
        Hidden
    }

    public sealed class ContentTypeFieldReference
    {
        public ContentTypeFieldReference(
            string name,
            ContentTypeFieldStatus status
        )
        {
            Name = name;
            Status = status;
        }

        public string Name { get; private set; }
        public ContentTypeFieldStatus Status { get; private set; }
    }
}

