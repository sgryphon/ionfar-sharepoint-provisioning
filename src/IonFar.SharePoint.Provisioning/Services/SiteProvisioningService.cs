using IonFar.SharePoint.Provisioning.Infrastructure;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using Microsoft.SharePoint.Client.Taxonomy;

namespace IonFar.SharePoint.Provisioning.Services
{
    public class SiteProvisioningService : ISiteProvisioningService
    {
        private readonly ClientContext _clientContext;
        private readonly ILogger _logger;

        public SiteProvisioningService(ClientContext clientContext, ILogger logger)
        {
            _clientContext = clientContext;
            _logger = logger;
        }

        public void CreateWeb(string url, string title)
        {
            _logger.Information("Creating web '{0}' with title '{1}'", url, title);

            var webCreationInformation = new WebCreationInformation
            {
                Title = title,
                Url = url,
                WebTemplate = Constants.WebTemplates.BlankSite
            };

            var newWeb = _clientContext.Web.Webs.Add(webCreationInformation);
            _clientContext.ExecuteQuery();

            var webNavigationSettings = new WebNavigationSettings(_clientContext, newWeb);
            webNavigationSettings.GlobalNavigation.Source = StandardNavigationSource.InheritFromParentWeb;

            var taxonomySession = TaxonomySession.GetTaxonomySession(_clientContext);
            webNavigationSettings.Update(taxonomySession);

            _clientContext.ExecuteQuery();
        }

        public void DeleteWeb(string url)
        {
            _logger.Warning("Deleting web '{0}'", url);

            var web = _clientContext.Site.OpenWeb(url);
            web.DeleteObject();

            _clientContext.ExecuteQuery();
        }
    }
}
