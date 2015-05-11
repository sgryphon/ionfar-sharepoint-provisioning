using System;
using System.Linq;
using IonFar.SharePoint.Provisioning.Infrastructure;
using Microsoft.SharePoint.Client;

namespace IonFar.SharePoint.Provisioning.Services
{
    public class FileUploadService : IFileUploadService
    {
        private readonly ClientContext _clientContext;
        private readonly ILogger _logger;

        private readonly string _apiUrl = string.Empty;
        
        public FileUploadService(ClientContext clientContext, ILogger logger)
        {
            _clientContext = clientContext;
            _logger = logger;
        }

        private void EnsureServerRelativeUrl()
        {
            if (!_clientContext.Web.IsPropertyAvailable("ServerRelativeUrl"))
            {
                _clientContext.Load(_clientContext.Web);
                _clientContext.ExecuteQuery();
            }
        }

        public void UploadFilesFromFolderToFolder(string sharepointFolderPath, string sourcePath, bool publishFiles, string fileSearchPattern = "*", bool includeSubdirectories = true)
        {
            sourcePath = sourcePath.TrimEnd('/', '\\');
            sharepointFolderPath = sharepointFolderPath.TrimEnd('/');

            EnsureServerRelativeUrl();

            var webRelativeUrl = _clientContext.Web.ServerRelativeUrl;
            var folderRelativeUrl = webRelativeUrl + "/" + sharepointFolderPath;

            var directory = new System.IO.DirectoryInfo(sourcePath);
            foreach (var file in directory.GetFiles(fileSearchPattern))
            {

                ReplaceWebUrl(file);
                ReplaceApiUrl(file);

                var folder = GetOrCreateFolder(folderRelativeUrl);

                var uploadedFile = folder.UploadFileWebDav(file.Name, file.FullName, overwriteIfExists: true);
                if (_logger != null)
                {
                    _logger.Information("Uploaded {0} to {1}", file.Name, folder.ServerRelativeUrl);
                }

                if (publishFiles)
                {
                    uploadedFile.Publish(string.Empty);
                    _clientContext.Load(uploadedFile);
                }

            }

            if (includeSubdirectories)
            {
                foreach (var sourcePathChild in directory.GetDirectories().Select(dir => dir.Name))
                {
                    UploadFilesFromFolderToFolder(sharepointFolderPath + "/" + sourcePathChild, sourcePath + "\\" + sourcePathChild, publishFiles, fileSearchPattern, true);
                }
            }

            _clientContext.ExecuteQuery();
        }

        private Folder GetOrCreateFolder(string folderRelativeUrl)
        {
            EnsureServerRelativeUrl();

            if (!folderRelativeUrl.StartsWith(_clientContext.Web.ServerRelativeUrl))
            {
                var msg = string.Format("You should not create a folder above the current Web root (web root: {0}, folder: {1})", 
                    _clientContext.Web.ServerRelativeUrl,
                    folderRelativeUrl);
                throw new Exception(msg);
            }

            var folder = _clientContext.Web.GetFolderByServerRelativeUrl(folderRelativeUrl);
            _clientContext.Load(folder);
            try
            {
                _clientContext.ExecuteQuery();
            }
            catch (ServerException)
            {
                var segments = folderRelativeUrl.Split(new [] {'/'}).ToList();
                var lastSegment = segments.Last();
                var parentFolderPath = string.Join("/", segments.Take(segments.Count() -1 ));
                var parentFolder = GetOrCreateFolder(parentFolderPath);
                folder = parentFolder.CreateFolder(lastSegment);
            }
            return folder;
        }

        public void UploadFilesFromFolderToListRootFolder(string sharePointListName, string folderPath, bool publishFiles, string fileSearchPattern = "*", bool includeSubdirectories = true)
        {
            EnsureServerRelativeUrl();

            var targetFolder = _clientContext.Web.Lists.GetByTitle(sharePointListName).RootFolder;
            _clientContext.Load(targetFolder);
            _clientContext.ExecuteQuery();

            var directory = new System.IO.DirectoryInfo(folderPath);
            foreach (var file in directory.GetFiles(fileSearchPattern, includeSubdirectories ? System.IO.SearchOption.AllDirectories : System.IO.SearchOption.TopDirectoryOnly))
            {
                _logger.Information("Uploading file: {0}", file.Name);
                ReplaceWebUrl(file);

                var uploadedFile = targetFolder.UploadFileWebDav(file.Name, file.FullName, true);
                if (publishFiles)
                {
                    uploadedFile.Publish(string.Empty);
                    _clientContext.Load(uploadedFile);
                }
                _clientContext.ExecuteQuery();
            }
        }

        public void UploadFilesFromFolderToList(string listName,string listFolderName, string localFolderPath, bool publishFiles, string fileSearchPattern = "*", bool includeSubdirectories = true)
        {
            EnsureServerRelativeUrl();

            var listFolders = _clientContext.Web.Lists.GetByTitle(listName).RootFolder.Folders;
            _clientContext.Load(listFolders);
            _clientContext.ExecuteQuery();

            Folder listFolder;
            if (!listFolders.Any(d => d.Name == listFolderName))
            {
                var rootFolder = _clientContext.Web.Lists.GetByTitle(listName).RootFolder;
                _clientContext.Load(rootFolder);
                _clientContext.ExecuteQuery();

                listFolder = rootFolder.CreateFolder(listFolderName);
            }
            else
            {
                listFolder = listFolders.GetByUrl(listFolderName);
                _clientContext.Load(listFolder);
                _clientContext.ExecuteQuery();
            }

            _logger.Information("Uploading files");

            var directory = new System.IO.DirectoryInfo(localFolderPath);
            foreach (var file in directory.GetFiles(fileSearchPattern, includeSubdirectories ? System.IO.SearchOption.AllDirectories : System.IO.SearchOption.TopDirectoryOnly))
            {
                _logger.Information("Uploading file: {0}", file.Name);
                ReplaceWebUrl(file);

                var uploadedFile = listFolder.UploadFileWebDav(file.Name, file.FullName, true);
                if (publishFiles)
                {
                    uploadedFile.Publish(string.Empty);
                    _clientContext.Load(uploadedFile);
                    _clientContext.ExecuteQuery();
                }
            }
        }

        /// <summary>
        /// Used to upload publishing pages to Pages library as UploadFileWebDav ( Office PNP ) doesn't work on that list
        /// </summary>
        /// <param name="sharepointFolderPath"></param>
        /// <param name="sourcePath"></param>
        /// <param name="publishFiles"></param>
        /// <param name="fileSearchPattern"></param>
        /// <param name="includeSubdirectories"></param>
        public void UploadFilesFromFolderToFolderWithoutWebDav(string sharepointFolderPath, string sourcePath, bool publishFiles, string fileSearchPattern = "*", bool includeSubdirectories = true)
        {
            EnsureServerRelativeUrl();

            sourcePath = sourcePath.TrimEnd('/', '\\');
            sharepointFolderPath = sharepointFolderPath.TrimEnd('/');

            _clientContext.Load(_clientContext.Web);
            _clientContext.ExecuteQuery();

            var webRelativeUrl = _clientContext.Web.ServerRelativeUrl;
            var folderRelativeUrl = webRelativeUrl + "/" + sharepointFolderPath;
            var folder = _clientContext.Web.GetFolderByServerRelativeUrl(folderRelativeUrl);
            _clientContext.Load(folder);
            _clientContext.Load(folder.Files);
            _clientContext.ExecuteQuery();

            var directory = new System.IO.DirectoryInfo(sourcePath);
            foreach (var file in directory.GetFiles(fileSearchPattern))
            {

                // not uploading .webpart files
                if (System.IO.Path.GetExtension(file.Name) == ".webpart") continue;

                ReplaceWebUrl(file);

                foreach (var f in folder.Files)
                {
                    if (file.Name.Equals(f.Name))
                    {
                        f.DeleteObject();
                        folder.Context.ExecuteQuery();
                        break;
                    }
                }

                var stream = System.IO.File.OpenRead(file.FullName);
                var newFileInfo = new FileCreationInformation
                {
                    ContentStream = stream,
                    Url = file.FullName.Replace(directory.FullName+"\\",string.Empty),
                    Overwrite = true
                };
                var uploadedFile = folder.Files.Add(newFileInfo);
                folder.Context.Load(uploadedFile);
                folder.Context.ExecuteQuery();

                if (_logger != null)
                {
                    _logger.Information("Uploaded {0} to {1}", file.Name, folder.ServerRelativeUrl);
                }

                if (publishFiles)
                {
                    uploadedFile.CheckIn(string.Empty,CheckinType.MajorCheckIn);
                    _clientContext.Load(uploadedFile);
                }

            }

            if (includeSubdirectories)
            {
                foreach (var sourcePathChild in directory.GetDirectories().Select(dir => dir.Name))
                {
                    UploadFilesFromFolderToFolder(sharepointFolderPath + "/" + sourcePathChild, sourcePath + "\\" + sourcePathChild, publishFiles, fileSearchPattern, true);
                }
            }

            _clientContext.ExecuteQuery();
        }

        private void ReplaceWebUrl(System.IO.FileInfo file)
        {
            var textFileExtensions = new[] {".css", ".js", ".html", ".webpart", ".aspx" };
            var webRelativeUrl = _clientContext.Web.ServerRelativeUrl;
            if (!_clientContext.Site.IsPropertyAvailable("ServerRelativeUrl"))
            {
                _clientContext.Load(_clientContext.Site, s => s.ServerRelativeUrl);
                _clientContext.ExecuteQuery();
            }
            var siteRelativeUrl = _clientContext.Site.ServerRelativeUrl;

            if (textFileExtensions.Contains(file.Extension))
            {
                var src = System.IO.File.ReadAllText(file.FullName);
                var dst = src.Replace("\"/SiteAssets", "\"{weburl}/SiteAssets")
                             .Replace("\"/_layouts", "\"{weburl}/_layouts")
                             .Replace("{weburl}", webRelativeUrl)
                             .Replace("{rooturl}", siteRelativeUrl);
                if (src != dst)
                {
                    System.IO.File.WriteAllText(file.FullName, dst);
                    _logger.Information("{{weburl}} token replaced in {0}", file.Name);
                }
            }
        }

        private void ReplaceApiUrl(System.IO.FileInfo file)
        {
            var src = System.IO.File.ReadAllText(file.FullName);
            var dst = src.Replace("{apiurl}", _apiUrl);
            if (src != dst)
            {
                System.IO.File.WriteAllText(file.FullName, dst);
                _logger.Information("{{apiurl}} token replaced in {0}", file.Name);
            }
        }
    }
}
