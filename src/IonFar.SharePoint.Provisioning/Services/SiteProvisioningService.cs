using IonFar.SharePoint.Provisioning.Infrastructure;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using Microsoft.SharePoint.Client.Taxonomy;

namespace IonFar.SharePoint.Provisioning.Services
{
    public class SiteProvisioningService : ISiteProvisioningService
    {
        private readonly ClientContext _clientContext;
        private readonly IProvisionLog _logger;

        /// <summary>
        /// Creates a new site provisioning service
        /// </summary>
        /// <param name="clientContext">Context to use</param>
        /// <param name="logger">(Optional) logger to use; if not specified defaults to TraceSource</param>
        public SiteProvisioningService(ClientContext clientContext, IProvisionLog logger = null)
        {
            _clientContext = clientContext;
            _logger = logger ?? new TraceProvisionLog();
        }

        /// <summary>
        /// Creates a new subweb with the specified properties, inheritting permissions and navigation.
        /// </summary>
        /// <param name="parentWeb">Parent under which to create the web, e.g. context.Site.RootWeb</param>
        /// <param name="leafUrl">URL component of the subweb, a single value</param>
        /// <param name="title">Title of the subweb</param>
        /// <param name="webTemplate">Template to use, e.g. "STS#0" for a team site; see the values in WebTemplates</param>
        /// <returns>The newly created Web</returns>
        public Web CreateWeb(Web parentWeb, string leafUrl, string title, string webTemplate)
        {
            return CreateWeb(parentWeb, leafUrl, title, webTemplate, description: null, inheritPermissions: true, inheritNavigation: true);
        }

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
        public Web CreateWeb(Web parentWeb, string leafUrl, string title, string webTemplate, string description, bool inheritPermissions, bool inheritNavigation)
        {
            _logger.Information("Creating web '{0}' with title '{1}'", leafUrl, title);

            var webCreationInformation = new WebCreationInformation
            {
                Title = title,
                Url = leafUrl,
                WebTemplate = webTemplate,
                Description = description,
                UseSamePermissionsAsParentSite = inheritPermissions                
            };

            var newWeb = _clientContext.Web.Webs.Add(webCreationInformation);
            _clientContext.ExecuteQuery();

            if (inheritNavigation)
            {
                var webNavigationSettings = new WebNavigationSettings(_clientContext, newWeb);
                webNavigationSettings.GlobalNavigation.Source = StandardNavigationSource.InheritFromParentWeb;

                var taxonomySession = TaxonomySession.GetTaxonomySession(_clientContext);
                webNavigationSettings.Update(taxonomySession);

                _clientContext.ExecuteQuery();
            }

            return newWeb;
        }

        /// <summary>
        /// Deletes the web at the specified server relative URL.
        /// </summary>
        /// <param name="serverRelativeUrl">URL of the site to delete</param>
        public void DeleteWeb(string serverRelativeUrl)
        {
            _logger.Warning("Deleting web '{0}'", serverRelativeUrl);

            var web = _clientContext.Site.OpenWeb(serverRelativeUrl);
            web.DeleteObject();

            _clientContext.ExecuteQuery();
        }
    }
}
