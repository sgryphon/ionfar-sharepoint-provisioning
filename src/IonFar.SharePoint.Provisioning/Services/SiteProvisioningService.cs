using IonFar.SharePoint.Provisioning.Infrastructure;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Linq;

namespace IonFar.SharePoint.Provisioning.Services
{
    /// <summary>
    /// Service for managing subwebs and features.
    /// </summary>
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
        /// Activates the given feature
        /// </summary>
        /// <param name="site">Site to acivate feature for</param>
        /// <param name="featureId">ID of the feature</param>
        public void ActivateFeature(Site site, Guid featureId)
        {
            ActivateFeature(site, featureId, FeatureDefinitionScope.Farm);
        }

        /// <summary>
        /// Activates the given feature
        /// </summary>
        /// <param name="web">Web to acivate feature for</param>
        /// <param name="featureId">ID of the feature</param>
        public void ActivateFeature(Web web, Guid featureId)
        {
            ActivateFeature(web, featureId, FeatureDefinitionScope.Farm);
        }

        /// <summary>
        /// Activates the given feature
        /// </summary>
        /// <param name="site">Site to acivate feature for</param>
        /// <param name="featureId">ID of the feature</param>
        /// <param name="scope">Scope of the definition (Farm for built in, Site for sandboxed)</param>
        public void ActivateFeature(Site site, Guid featureId, FeatureDefinitionScope definitionScope)
        {
            _clientContext.Load(site);
            _clientContext.Load(site.Features);
            _clientContext.ExecuteQuery();
            _logger.Information("Activating feature {0} in site '{1}", featureId, site.ServerRelativeUrl);
            InternalActivateFeature(site.Features, featureId, definitionScope, true);
        }

        /// <summary>
        /// Activates the given feature
        /// </summary>
        /// <param name="web">Web to acivate feature for</param>
        /// <param name="featureId">ID of the feature</param>
        /// <param name="scope">Scope of the definition (Farm for built in, Site for sandboxed)</param>
        public void ActivateFeature(Web web, Guid featureId, FeatureDefinitionScope definitionScope)
        {
            _clientContext.Load(web);
            _clientContext.Load(web.Features);
            _clientContext.ExecuteQuery();
            _logger.Information("Activating feature {0} in web '{1}", featureId, web.ServerRelativeUrl);
            InternalActivateFeature(web.Features, featureId, definitionScope, true);
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

            var newWeb = parentWeb.Webs.Add(webCreationInformation);
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

        /// <summary>
        /// Creates or updates a site collection ScriptLink reference to a script file, doing nothing if the ScriptLink already exists with the specified values
        /// </summary>
        /// <param name="name">Key to identify the ScriptLink</param>
        /// <param name="scriptPrefixedUrl">URL of the script; may use '~sitecollection/' or '~site/' prefix.</param>
        /// <param name="sequence">Determines the order the ScriptLink is rendered in</param>
        /// <returns>The UserCustomAction representing the ScriptLink</returns>
        public UserCustomAction EnsureSiteScriptLink(string name, string scriptPrefixedUrl, int sequence)
        {
            var site = _clientContext.Site;
            if (!site.IsObjectPropertyInstantiated("UserCustomActions"))
            {
                _clientContext.Load(site.UserCustomActions, collection => collection.Include(ca => ca.Name));
                _clientContext.ExecuteQuery();
            }
            var action = site.UserCustomActions.FirstOrDefault(ca => string.Equals(ca.Name, name, StringComparison.InvariantCultureIgnoreCase));
            if (action == null)
            {
                action = site.UserCustomActions.Add();
                action.Location = "ScriptLink";
                action.Name = name;
                action.ScriptSrc = scriptPrefixedUrl;
                action.Sequence = sequence;
                action.Update();
                _logger.Information("Adding ScriptLink '{0}'='{1}'", name, scriptPrefixedUrl);
                _clientContext.ExecuteQuery();
            }
            else
            {
                _clientContext.Load(action);
                _clientContext.ExecuteQuery();
                bool changed = false;
                if (action.Location != "ScriptLink")
                {
                    action.Location = "ScriptLink";
                    changed = true;
                }
                if (action.Name != name)
                {
                    action.Name = name;
                    changed = true;
                }
                if (action.ScriptSrc != scriptPrefixedUrl)
                {
                    action.ScriptSrc = scriptPrefixedUrl;
                    changed = true;
                }
                if (action.Sequence != sequence)
                {
                    action.Sequence = sequence;
                    changed = true;
                }
                if (changed)
                {
                    _logger.Information("Updating ScriptLink '{0}'='{1}'", name, scriptPrefixedUrl);
                    action.Update();
                    _clientContext.ExecuteQuery();
                }
                else
                {
                    _logger.Information("No change for ScriptLink '{0}'", name);
                }
            }
            return action;
        }

        private void InternalActivateFeature(FeatureCollection featureCollection, Guid featureId, FeatureDefinitionScope definitionScope, bool activate)
        {
            if (activate)
            {
                featureCollection.Add(featureId, true, definitionScope);
                _clientContext.ExecuteQuery();
            }
            else
            {
                throw new NotImplementedException("Feature deactivation not supported");
            }
        }
    }
}
