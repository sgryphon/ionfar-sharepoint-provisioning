namespace IonFar.SharePoint.Provisioning.Services
{
    public interface ISiteProvisioningService
    {
        void CreateWeb(string url, string title);
        void DeleteWeb(string url);
    }
}