using Microsoft.SharePoint.Client;

namespace IonFar.SharePoint.Provisioning.Services
{
    public interface IListProvisioningService
    {
        void CreateList(string listTitle, ListTemplateType listTemplateType, string parentWeb, bool enableContentTypes, bool enableModeration = false);
        void DeleteList(string listName, string parentWeb);
        void EnsureAssetsLibrary(string parentWeb);
        void AddContentTypeToList(string listName, string parentWeb, string contentTypeId);
        void DeleteContentTypeFromList(string listName, string parentWeb, string contentTypeName);
        void AddFieldsToDefaultView(string listName, string parentWeb, string[] fieldNames);
        void RenameField(string listName, string parentWeb, string originalFieldName, string newFieldName);
    }
}