using Microsoft.SharePoint.Client;

namespace IonFar.SharePoint.Provisioning.Services
{
    public interface IListProvisioningService
    {
        void AddContentTypeToList(string listName, string parentWeb, string contentTypeId);
        List CreateList(ListDescriptor listDescriptor);
        void EnsureSiteAssetsLibrary(string parentWeb);
        void EnsureSitePagesLibrary(string parentWeb);
    }
}