using IonFar.SharePoint.Provisioning;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Microsoft.SharePoint.Client
{
    static class Extensions
    {
        const string REGEX_INVALID_FILE_NAME_CHARS = @"[<>:;*?/\\|""&%\t\r\n]";

        /// <summary>
        /// Uploads a file to the specified folder by saving the binary directly (via webdav).
        /// </summary>
        /// <param name="folder">Folder to upload file to.</param>
        /// <param name="fileName"></param>
        /// <param name="localFilePath">Location of the file to be uploaded.</param>
        /// <param name="overwriteIfExists">true (default) to overwite existing files</param>
        /// <returns>The uploaded File, so that additional operations (such as setting properties) can be done.</returns>
        public static File UploadFileWebDav(this Folder folder, string fileName, string localFilePath, bool overwriteIfExists)
        {
            if (folder == null)
            {
                throw new ArgumentNullException("folder");
            }

            if (localFilePath == null)
            {
                throw new ArgumentNullException("localFilePath");
            }

            if (!System.IO.File.Exists(localFilePath))
            {
                throw new FileNotFoundException("Local file was not found.", localFilePath);
            }

            using (var stream = System.IO.File.OpenRead(localFilePath))
                return folder.UploadFileWebDav(fileName, stream, overwriteIfExists);
        }

        /// <summary>
        /// Uploads a file to the specified folder by saving the binary directly (via webdav).
        /// Note: this method does not work using app only token.
        /// </summary>
        /// <param name="folder">Folder to upload file to.</param>
        /// <param name="fileName">Location of the file to be uploaded.</param>
        /// <param name="stream"></param>
        /// <param name="overwriteIfExists">true (default) to overwite existing files</param>
        /// <returns>The uploaded File, so that additional operations (such as setting properties) can be done.</returns>
        public static File UploadFileWebDav(this Folder folder, string fileName, Stream stream, bool overwriteIfExists)
        {
            if (fileName == null)
            {
                throw new ArgumentNullException("fileName");
            }

            if (stream == null)
            {
                throw new ArgumentNullException("stream");
            }

            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentException("Destination_file_name_is_required_", "fileName");
            }

            if (Regex.IsMatch(fileName, REGEX_INVALID_FILE_NAME_CHARS))
            {
                throw new ArgumentException("The_argument_must_be_a_single_file_name_and_cannot_contain_path_characters_", "fileName");
            }

            var serverRelativeUrl = UrlUtility.Combine(folder.ServerRelativeUrl, fileName);

            // Create uploadContext to get a proper ClientContext instead of a ClientRuntimeContext
            using (var uploadContext = folder.Context.Clone(folder.Context.Url))
            {
                //Log.Debug(Constants.LOGGING_SOURCE, "Save binary direct (via webdav) to '{0}'", serverRelativeUrl);
                File.SaveBinaryDirect(uploadContext, serverRelativeUrl, stream, overwriteIfExists);
                uploadContext.ExecuteQueryRetry();
            }

            var file = folder.Files.GetByUrl(serverRelativeUrl);
            folder.Context.Load(file);
            folder.Context.ExecuteQueryRetry();

            return file;
        }

        /// <summary>
        /// Clones a ClientContext object while "taking over" the security context of the existing ClientContext instance
        /// </summary>
        /// <param name="clientContext">ClientContext to be cloned</param>
        /// <param name="siteUrl">Site url to be used for cloned ClientContext</param>
        /// <returns>A ClientContext object created for the passed site url</returns>
        public static ClientContext Clone(this ClientRuntimeContext clientContext, string siteUrl)
        {
            if (string.IsNullOrWhiteSpace(siteUrl))
            {
                throw new ArgumentException("Url_of_the_site_is_required_", "siteUrl");
            }

            return clientContext.Clone(new Uri(siteUrl));
        }

        /// <summary>
        /// Clones a ClientContext object while "taking over" the security context of the existing ClientContext instance
        /// </summary>
        /// <param name="clientContext">ClientContext to be cloned</param>
        /// <param name="siteUrl">Site url to be used for cloned ClientContext</param>
        /// <returns>A ClientContext object created for the passed site url</returns>
        public static ClientContext Clone(this ClientRuntimeContext clientContext, Uri siteUrl)
        {
            if (siteUrl == null)
            {
                throw new ArgumentException("Url_of_the_site_is_required_", "siteUrl");
            }

            ClientContext clonedClientContext = new ClientContext(siteUrl);
            clonedClientContext.AuthenticationMode = clientContext.AuthenticationMode;

            // In case of using networkcredentials in on premises or SharePointOnlineCredentials in Office 365
            if (clientContext.Credentials != null)
            {
                clonedClientContext.Credentials = clientContext.Credentials;
            }
            else
            {
                //Take over the form digest handling setting
                clonedClientContext.FormDigestHandlingEnabled = (clientContext as ClientContext).FormDigestHandlingEnabled;

                // In case of app only or SAML
                clonedClientContext.ExecutingWebRequest += delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    // Call the ExecutingWebRequest delegate method from the original ClientContext object, but pass along the webRequestEventArgs of 
                    // the new delegate method
                    MethodInfo methodInfo = clientContext.GetType().GetMethod("OnExecutingWebRequest", BindingFlags.Instance | BindingFlags.NonPublic);
                    object[] parametersArray = new object[] { webRequestEventArgs };
                    methodInfo.Invoke(clientContext, parametersArray);
                };
            }

            return clonedClientContext;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        public static void ExecuteQueryRetry(this ClientRuntimeContext clientContext, int retryCount = 10, int delay = 500)
        {
            ExecuteQueryImplementation(clientContext, retryCount, delay);
        }

        private static void ExecuteQueryImplementation(ClientRuntimeContext clientContext, int retryCount = 10, int delay = 500)
        {
            int retryAttempts = 0;
            int backoffInterval = delay;
            if (retryCount <= 0)
                throw new ArgumentException("Provide a retry count greater than zero.");

            if (delay <= 0)
                throw new ArgumentException("Provide a delay greater than zero.");

            // Do while retry attempt is less than retry count
            while (retryAttempts < retryCount)
            {
                try
                {
                    clientContext.ExecuteQuery();
                    return;

                }
                catch (WebException wex)
                {
                    var response = wex.Response as HttpWebResponse;
                    // Check if request was throttled - http status code 429
                    // Check is request failed due to server unavailable - http status code 503
                    if (response != null && (response.StatusCode == (HttpStatusCode)429 || response.StatusCode == (HttpStatusCode)503))
                    {
                        Debug.WriteLine("CSOM request frequency exceeded usage limits. Sleeping for {0} seconds before retrying.", backoffInterval);

                        //Add delay for retry
                        Thread.Sleep(backoffInterval);

                        //Add to retry count and increase delay.
                        retryAttempts++;
                        backoffInterval = backoffInterval * 2;
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            //throw new MaximumRetryAttemptedException(string.Format("Maximum retry attempts {0}, has be attempted.", retryCount));
            throw new Exception(string.Format("Maximum retry attempts {0}, has be attempted.", retryCount));
        }

        /// <summary>
        /// Creates a folder with the given name.
        /// </summary>
        /// <param name="parentFolder">Parent folder to create under</param>
        /// <param name="folderName">Folder name to retrieve or create</param>
        /// <returns>The newly created folder</returns>
        /// <remarks>
        /// <para>
        /// Note that this only checks one level of folder (the Folders collection) and cannot accept a name with path characters.
        /// </para>
        /// <example>
        ///     var folder = list.RootFolder.CreateFolder("new-folder");
        /// </example>
        /// </remarks>
        public static Folder CreateFolder(this Folder parentFolder, string folderName)
        {
            // TODO: Check for any other illegal characters in SharePoint
            if (folderName.Contains('/') || folderName.Contains('\\'))
            {
                throw new ArgumentException("The_argument_must_be_a_single_folder_name_and_cannot_contain_path_characters_", "folderName");
            }

            var folderCollection = parentFolder.Folders;
            var folder = CreateFolderImplementation(folderCollection, folderName);
            return folder;
        }

        private static Folder CreateFolderImplementation(FolderCollection folderCollection, string folderName)
        {
            var newFolder = folderCollection.Add(folderName);
            folderCollection.Context.Load(newFolder);
            folderCollection.Context.ExecuteQueryRetry();

            return newFolder;
        }

        /// <summary>
        /// Sets a key/value pair in the web property bag
        /// </summary>
        /// <param name="web">Web that will hold the property bag entry</param>
        /// <param name="key">Key for the property bag entry</param>
        /// <param name="value">String value for the property bag entry</param>
        public static void SetPropertyBagValue(this Web web, string key, string value)
        {
            SetPropertyBagValueInternal(web, key, value);
        }

        /// <summary>
        /// Sets a key/value pair in the web property bag
        /// </summary>
        /// <param name="web">Web that will hold the property bag entry</param>
        /// <param name="key">Key for the property bag entry</param>
        /// <param name="value">Value for the property bag entry</param>
        private static void SetPropertyBagValueInternal(Web web, string key, object value)
        {
            var props = web.AllProperties;

            // Get the value, if the web properties are already loaded
            if (props.FieldValues.Count > 0)
            {
                props[key] = value;
            }
            else

            {
                // Load the web properties
                web.Context.Load(props);
                web.Context.ExecuteQueryRetry();

                props[key] = value;
            }

            web.Update();
            web.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Get string typed property bag value. If does not contain, returns given default value.
        /// </summary>
        /// <param name="web">Web to read the property bag value from</param>
        /// <param name="key">Key of the property bag entry to return</param>
        /// <param name="defaultValue"></param>
        /// <returns>Value of the property bag entry as string</returns>
        public static string GetPropertyBagValueString(this Web web, string key, string defaultValue)
        {
            object value = GetPropertyBagValueInternal(web, key);
            if (value != null)
            {
                return (string)value;
            }
            else
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// Type independent implementation of the property getter.
        /// </summary>
        /// <param name="web">Web to read the property bag value from</param>
        /// <param name="key">Key of the property bag entry to return</param>
        /// <returns>Value of the property bag entry</returns>
        private static object GetPropertyBagValueInternal(Web web, string key)
        {
            var props = web.AllProperties;
            web.Context.Load(props);
            web.Context.ExecuteQueryRetry();
            if (props.FieldValues.ContainsKey(key))
            {
                return props.FieldValues[key];
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Check if a folder exists with the specified path (relative to the web), and if not creates it (inside a list if necessary)
        /// </summary>
        /// <param name="web">Web to check for the specified folder</param>
        /// <param name="webRelativeUrl">Path to the folder, relative to the web site</param>
        /// <returns>The existing or newly created folder</returns>
        /// <remarks>
        /// <para>
        /// If the specified path is inside an existing list, then the folder is created inside that list.
        /// </para>
        /// <para>
        /// Any existing folders are traversed, and then any remaining parts of the path are created as new folders.
        /// </para>
        /// </remarks>
        public static Folder EnsureFolderPath(this Web web, string webRelativeUrl)
        {
            if (webRelativeUrl == null) { throw new ArgumentNullException("webRelativeUrl"); }
            if (string.IsNullOrWhiteSpace(webRelativeUrl)) { throw new ArgumentException("Folder_URL_is_required_", "webRelativeUrl"); }

            // Check if folder exists
            if (!web.IsObjectPropertyInstantiated("ServerRelativeUrl"))
            {
                web.Context.Load(web, w => w.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();
            }
            var folderServerRelativeUrl = web.ServerRelativeUrl + (web.ServerRelativeUrl.EndsWith("/") ? "" : "/") + webRelativeUrl;

            // Check if folder is inside a list
            var listCollection = web.Lists;
            web.Context.Load(listCollection, lc => lc.Include(l => l.RootFolder.ServerRelativeUrl));
            web.Context.ExecuteQueryRetry();

            List containingList = null;
            foreach (var list in listCollection)
            {
                if (folderServerRelativeUrl.StartsWith(list.RootFolder.ServerRelativeUrl, StringComparison.InvariantCultureIgnoreCase))
                {
                    containingList = list;
                    break;
                }
            }

            // Start either at the root of the list or web
            string locationType = null;
            string rootUrl = null;
            Folder currentFolder = null;
            if (containingList == null)
            {
                locationType = "Web";
                currentFolder = web.RootFolder;
                web.Context.Load(currentFolder, f => f.ServerRelativeUrl);
                web.Context.ExecuteQueryRetry();
            }
            else
            {
                locationType = "List";
                currentFolder = containingList.RootFolder;
            }

            rootUrl = currentFolder.ServerRelativeUrl;

            // Get remaining parts of the path and split
            var folderRootRelativeUrl = folderServerRelativeUrl.Substring(currentFolder.ServerRelativeUrl.Length);
            var childFolderNames = folderRootRelativeUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var currentCount = 0;

            foreach (var folderName in childFolderNames)
            {
                currentCount++;

                // Find next part of the path
                var folderCollection = currentFolder.Folders;
                folderCollection.Context.Load(folderCollection);
                folderCollection.Context.ExecuteQueryRetry();
                Folder nextFolder = null;
                foreach (Folder existingFolder in folderCollection)
                {
                    if (string.Equals(existingFolder.Name, folderName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        nextFolder = existingFolder;
                        break;
                    }
                }

                // Or create it
                if (nextFolder == null)
                {
                    var createPath = string.Join("/", childFolderNames, 0, currentCount);
                    //Log.Info(Constants.LOGGING_SOURCE, CoreResources.FileFolderExtensions_CreateFolder0Under12, createPath, locationType, rootUrl);

                    nextFolder = folderCollection.Add(folderName);
                    folderCollection.Context.Load(nextFolder);
                    folderCollection.Context.ExecuteQueryRetry();
                }

                currentFolder = nextFolder;
            }

            return currentFolder;
        }

    }

}
