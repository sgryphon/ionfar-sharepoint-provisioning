using Microsoft.SharePoint.Client;

namespace IonFar.SharePoint.Provisioning.Services
{
    public interface ISiteProvisioningService
    {
        /// <summary>
        /// Creates a new subweb with the specified properties.
        /// </summary>
        /// <param name="parentWeb">Parent under which to create the web, e.g. context.Site.RootWeb</param>
        /// <param name="leafUrl">URL component of the subweb, a single value</param>
        /// <param name="title">Title of the subweb</param>
        /// <param name="webTemplate">Template to use, e.g. "STS#0" for a team site; see the values in WebTemplates</param>
        /// <returns>The newly created Web</returns>
        Web CreateWeb(Web parentWeb, string leafUrl, string title, string webTemplate);

        /// <summary>
        /// Creates a new subweb with the specified properties.
        /// </summary>
        /// <param name="parentWeb">Parent under which to create the web, e.g. context.Site.RootWeb</param>
        /// <param name="leafUrl">URL component of the subweb, a single value</param>
        /// <param name="title">Title of the subweb</param>
        /// <param name="webTemplate">Template to use, e.g. "STS#0" for a team site; see the values in WebTemplates</param>
        /// <param name="description">Description of the subweb</param>
        /// <param name="inheritPermissions">true to inherit permissions; false for unique permissions</param>
        /// <param name="inheritNavigation">true to inherit navigation; false to not</param>
        /// <returns>The newly created Web</returns>
        Web CreateWeb(Web parentWeb, string leafUrl, string title, string webTemplate, string description, bool inheritPermissions, bool inheritNavigation);

        void DeleteWeb(string url);
    }
}