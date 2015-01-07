namespace IonFar.SharePoint.Provisioning.Services
{
    public interface IContentTypeProvisioningService
    {
        void CreateContentType(string contentTypeName, string contentTypeDescription, string contentTypeGroup, string contentTypeId);
        void DeleteContentType(string contentTypeId);
        void AddFieldLinkToContentType(string contentTypeId, string fieldName);
    }
}