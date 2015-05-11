using IonFar.SharePoint.Provisioning.Services;
using System.Collections.Generic;

namespace IonFar.SharePoint.Provisioning.Services
{
    public interface IContentTypeProvisioningService
    {
        void Provision(IEnumerable<ContentTypeDescriptor> contentTypeDescriptor);
        void Provision(ContentTypeDescriptor contentTypeDescriptor);

        void CreateContentType(string contentTypeName, string contentTypeDescription, string contentTypeGroup, string contentTypeId);
        void DeleteContentType(string contentTypeId);
        void AddFieldLinkToContentType(string contentTypeId, string fieldName);
    }
}