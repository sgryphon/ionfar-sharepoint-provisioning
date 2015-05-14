using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using IonFar.SharePoint.Provisioning.Infrastructure;
using System.Net;
//using Serilog;

namespace IonFar.SharePoint.Provisioning.Services
{
    public class DirectorySync
    {
        private const string PropertyBagKey = "SPSync.DirectorySync";
        private readonly ICredentials _credentials;
        private IDictionary<string, string> _filePathToChecksum;
        private readonly ContentTransformer _contentTransformer;
        private IProvisionLog _logger;
        private readonly Uri _sharepointServer;
        private readonly object _thelock = new object();

        public DirectorySync(Uri sharepointServer, Uri apiServer, ICredentials credentials, bool reset)
        {
            _sharepointServer = sharepointServer;
            _credentials = credentials;
            // Can't load during constructor as Logger not yet injected.
            _filePathToChecksum = reset ? new Dictionary<string, string>() : null;
            _contentTransformer = new ContentTransformer(sharepointServer, sharepointServer, apiServer);
        }

        public IProvisionLog Logger
        {
            set { _logger = value; }
        }

        public void Execute(string localDirectory, string serverDirectory, bool isAsynchronous)
        {
            if (!Directory.Exists(localDirectory)) throw new ArgumentException(localDirectory + " not found");
            if (_filePathToChecksum == null)
            {
                _filePathToChecksum = LoadServerChecksums();
            }

            serverDirectory += serverDirectory.EndsWith("/") ? "" : "/";
            localDirectory += localDirectory.EndsWith("\\") ? "" : "\\";

            var paths = Directory.EnumerateFiles(localDirectory, "*.*", SearchOption.AllDirectories);
            var tasks = isAsynchronous ? new List<Task>() : null;
            foreach (var path in paths)
            {
                var serverRelativePath = GetRelativeServerPath(localDirectory, serverDirectory, path);
                if (LocalChecksum(path) != ServerChecksum(serverRelativePath))
                {
                    HandleFileChanged(localDirectory, serverDirectory, tasks, path);
                }
                else
                {
                    if (_logger != null)
                    {
                        _logger.Information("SPSync {0} {1}", "Up-to-date", path);
                    }
                }
            }

            if (tasks != null) Task.WaitAll(tasks.ToArray());
            SaveServerChecksums(_filePathToChecksum);
        }

        private void HandleFileChanged(string localDirectory, string serverDirectory, List<Task> tasks, string path)
        {
            var localUri = new Uri(Path.GetFullPath(path));
            var localDirUri = new Uri(Path.GetFullPath(localDirectory));

            var relativePath = localDirUri
                .MakeRelativeUri(localUri);
            var relativeServerPath = serverDirectory + relativePath;

            if (tasks == null)
            {
                Upload(Path.GetFullPath(path), relativeServerPath);
                UpdateServerChecksum(path, relativeServerPath);
            }
            else
            {
                var t = UploadAndUpdateServerChecksumAsync(Path.GetFullPath(path), relativeServerPath);
                tasks.Add(t);
            }
        }

        private async Task UploadAndUpdateServerChecksumAsync(string path, string serverPath)
        {
            await UploadAsync(path, serverPath);
            UpdateServerChecksum(path, serverPath);
        }

        private IDictionary<string, string> LoadServerChecksums()
        {
            try
            {
                using (var spContext = new ClientContext(_sharepointServer))
                {
                    spContext.Credentials = _credentials;
                    var json = spContext.Web.GetPropertyBagValueString(PropertyBagKey, "[]");
                    var list = JsonConvert.DeserializeObject<IList<KeyValuePair<string, string>>>(json);
                    var dict = list.ToDictionary(kv => kv.Key, kv => kv.Value);
                    if (dict.Count == 0)
                    {
                        if (_logger != null)
                        {
                            _logger.Warning(PropertyBagKey + " not found in property bag. Prepare for all static files to be deployed.");
                        }
                    }
                    else
                    {
                        if (_logger != null)
                        {
                            _logger.Information("Incremental Deploy enabled using " + PropertyBagKey + " in property bag");
                        }
                    }
                    return dict;
                }
            }
            catch (Exception exception)
            {
                if (_logger != null)
                {
                    _logger.Error("Incremental Deploy disabled because error reading '{0}' from property bag: {1}", PropertyBagKey, exception);
                }
                return new Dictionary<string, string>();
            }
        }

        private void SaveServerChecksums(IDictionary<string, string> dict)
        {
            lock (_thelock)
            {
                using (var spContext = new ClientContext(_sharepointServer))
                {
                    spContext.Credentials = _credentials;
                    var json = JsonConvert.SerializeObject(new List<KeyValuePair<string, string>>(dict));
                    spContext.Web.SetPropertyBagValue(PropertyBagKey, json);
                }
            }
        }

        private void UpdateServerChecksum(string path, string serverRelativePath)
        {
            _filePathToChecksum[serverRelativePath] = LocalChecksum(path);
            if (_logger != null)
            {
                _logger.Information("SPSync {0} {1}", "Uploaded", path);
            }
        }

        private Task UploadAsync(string path, string serverPath)
        {
            return Task.Run(() => Upload(path, serverPath));
        }

        private void Upload(string path, string serverPath)
        {
            using (var spContext = new ClientContext(_sharepointServer))
            {
                spContext.Credentials = _credentials;

                var directoryName = Path.GetDirectoryName(serverPath);
                if (directoryName != null)
                {
                    var relativeUrl = directoryName.Replace("\\", "/");

                    var folder = spContext.Web.EnsureFolderPath(relativeUrl);
                    var fileName = Path.GetFileName(path);

                    var textFileExtensions = new[] {".css", ".js", ".html", ".webpart", ".aspx"};
                    if (textFileExtensions.Contains(Path.GetExtension(path)))
                    {
                        var contents = _contentTransformer.Fixup(path);
                        using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(contents)))
                        {
                            folder.UploadFileWebDav(fileName, stream, true);
                        }
                    }
                    else
                    {
                        folder.UploadFileWebDav(fileName, path, true);
                    }
                }
            }
        }

        private string LocalChecksum(string path)
        {
            var fixedUpContent = _contentTransformer.Fixup(path);
            using (var stream = new MemoryStream())
            using (var writer = new StreamWriter(stream))
            {
                writer.Write(fixedUpContent);
                writer.Flush();
                stream.Position = 0;

                var sha1 = SHA1.Create();
                var hash = sha1.ComputeHash(stream);
                return BitConverter.ToString(hash);
            }
        }

        private string ServerChecksum(string serverRelativePath)
        {
            // forces file to be uploaded
            //if (path.ToLower().Contains("angular")) return "";

            //path = Path.GetFullPath(path);
            return _filePathToChecksum.ContainsKey(serverRelativePath) ? _filePathToChecksum[serverRelativePath] : "";
        }

        private static string GetRelativeServerPath(string localDirectory, string serverDirectory, string path)
        {
            var localUri = new Uri(Path.GetFullPath(path));
            var localDirUri = new Uri(Path.GetFullPath(localDirectory));

            var relativePath = localDirUri
                .MakeRelativeUri(localUri);
            var relativeServerPath = serverDirectory + relativePath;
            return relativeServerPath;
        }
    }
}