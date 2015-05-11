namespace IonFar.SharePoint.Provisioning.Services
{
    public interface IFileUploadService
    {
        void UploadFilesFromFolderToFolder(string sharepointFolderPath, string sourcePath, bool publishFiles, string fileSearchPattern, bool includeSubdirectories);
        void UploadFilesFromFolderToListRootFolder(string sharePointListName, string folderPath, bool publishFiles, string fileSearchPattern, bool includeSubdirectories);
        void UploadFilesFromFolderToList(string listName,string listFolderName, string localFolderPath, bool publishFiles, string fileSearchPattern, bool includeSubdirectories);
        void UploadFilesFromFolderToFolderWithoutWebDav(string sharepointFolderPath, string sourcePath, bool publishFiles, string fileSearchPattern, bool includeSubdirectories);
    }
}