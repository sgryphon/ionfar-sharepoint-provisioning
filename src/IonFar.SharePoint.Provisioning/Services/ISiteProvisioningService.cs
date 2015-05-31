using Microsoft.SharePoint.Client;
using System;

namespace IonFar.SharePoint.Provisioning.Services
{
    public interface ISiteProvisioningService
    {
        /// <summary>
        /// Activates the given feature
        /// </summary>
        /// <param name="site">Site to acivate feature for</param>
        /// <param name="featureId">ID of the feature</param>
        void ActivateFeature(Site site, Guid featureId);

        /// <summary>
        /// Activates the given feature
        /// </summary>
        /// <param name="web">Web to acivate feature for</param>
        /// <param name="featureId">ID of the feature</param>
        void ActivateFeature(Web web, Guid featureId);

        /// <summary>
        /// Activates the given feature
        /// </summary>
        /// <param name="site">Site to acivate feature for</param>
        /// <param name="featureId">ID of the feature</param>
        /// <param name="scope">Scope of the definition (Farm for built in, Site for sandboxed)</param>
        void ActivateFeature(Site site, Guid featureId, FeatureDefinitionScope definitionScope);

        /// <summary>
        /// Activates the given feature
        /// </summary>
        /// <param name="web">Web to acivate feature for</param>
        /// <param name="featureId">ID of the feature</param>
        /// <param name="scope">Scope of the definition (Farm for built in, Site for sandboxed)</param>
        void ActivateFeature(Web web, Guid featureId, FeatureDefinitionScope definitionScope);

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

        /// <summary>
        /// Creates or updates a site collection ScriptLink reference to a script file, doing nothing if the ScriptLink already exists with the specified values
        /// </summary>
        /// <param name="name">Key to identify the ScriptLink</param>
        /// <param name="scriptPrefixedUrl">URL of the script; may use '~sitecollection/' or '~site/' prefix.</param>
        /// <param name="sequence">Determines the order the ScriptLink is rendered in</param>
        /// <returns>The UserCustomAction representing the ScriptLink</returns>
        UserCustomAction EnsureSiteScriptLink(string name, string scriptPrefixedUrl, int sequence);

    }
}