using System.Linq;
using IonFar.SharePoint.Provisioning.Infrastructure;
using Microsoft.SharePoint.Client;

namespace IonFar.SharePoint.Provisioning.Services
{
    public class ListProvisioningService : IListProvisioningService
    {
        private readonly ClientContext _clientContext;
        private readonly IProvisionLog _logger;

        public ListProvisioningService(ClientContext clientContext, IProvisionLog logger)
        {
            _clientContext = clientContext;
            _logger = logger;
        }

        public void CreateList(string listTitle, ListTemplateType listTemplateType, string parentWeb, bool enableContentTypes, bool enableModeration = false)
        {
            _logger.Information("Creating list '{0}' at '{1}'", listTitle, parentWeb);

            var hostWeb = string.IsNullOrWhiteSpace(parentWeb)
                ? _clientContext.Site.RootWeb
                : _clientContext.Site.OpenWeb(parentWeb);

            var lists = hostWeb.Lists;
            var listCreationInformation = new ListCreationInformation
            {
                QuickLaunchOption = QuickLaunchOptions.On,
                Title = listTitle,
                TemplateType = (int)listTemplateType
            };

            var list = lists.Add(listCreationInformation);
            list.Update();

            _clientContext.ExecuteQuery();

            list.ContentTypesEnabled = enableContentTypes;
            list.EnableModeration = enableModeration;
            list.Update();
            _clientContext.ExecuteQuery();
        }

        public void DeleteList(string listName, string parentWeb)
        {
            _logger.Warning("Deleting list '{0}'", listName);

            var hostWeb = string.IsNullOrWhiteSpace(parentWeb)
                ? _clientContext.Site.RootWeb
                : _clientContext.Site.OpenWeb(parentWeb);

            var listToDelete = hostWeb.Lists.GetByTitle(listName);
            listToDelete.DeleteObject();

            _clientContext.ExecuteQuery();

            _logger.Information("List '{0}' deleted", listName);
        }

        public void EnsureAssetsLibrary(string parentWeb)
        {
            _logger.Information("Creating Assets Library at '{0}'", parentWeb);

            var hostWeb = string.IsNullOrWhiteSpace(parentWeb)
                ? _clientContext.Site.RootWeb
                : _clientContext.Site.OpenWeb(parentWeb);

            var lists = hostWeb.Lists;
            _clientContext.Load(lists);

            lists.EnsureSiteAssetsLibrary();
        }

        public void AddContentTypeToList(string listName, string parentWeb, string contentTypeId)
        {
            _logger.Information("Adding Content Type Id '{0}' to list '{1}' at web '{2}'", contentTypeId, listName, parentWeb);

            var hostWeb = string.IsNullOrWhiteSpace(parentWeb)
                ? _clientContext.Site.RootWeb
                : _clientContext.Site.OpenWeb(parentWeb);

            var lists = hostWeb.Lists;

            _clientContext.Load(lists);
            _clientContext.ExecuteQuery();

            var list = lists.FirstOrDefault(l => l.Title == listName);
            var contentType = _clientContext.Site.RootWeb.ContentTypes.GetById(contentTypeId);

            _clientContext.Load(list);
            _clientContext.Load(contentType);

            list.ContentTypes.AddExistingContentType(contentType);

            _clientContext.ExecuteQuery();
        }

        public void DeleteContentTypeFromList(string listName, string parentWeb, string contentTypeName)
        {
            _logger.Warning("Deleting content type '{0}' from list '{1}'", contentTypeName, listName);

            var hostWeb = string.IsNullOrWhiteSpace(parentWeb)
                ? _clientContext.Site.RootWeb
                : _clientContext.Site.OpenWeb(parentWeb);

            var lists = hostWeb.Lists;

            _clientContext.Load(lists);
            _clientContext.ExecuteQuery();

            var list = lists.FirstOrDefault(l => l.Title == listName);
            var contentTypes = list.ContentTypes;

            _clientContext.Load(list);
            _clientContext.Load(contentTypes);
            _clientContext.ExecuteQuery();

            var contentType = list.ContentTypes.FirstOrDefault(c => c.Name == contentTypeName);
            contentType.DeleteObject();

            _clientContext.ExecuteQuery();
        }

        public void AddFieldsToDefaultView(string listName, string parentWeb, string[] fieldNames)
        {
            _logger.Information("Adding fields to default view of '{0}'", listName);

            var hostWeb = string.IsNullOrWhiteSpace(parentWeb)
                ? _clientContext.Site.RootWeb
                : _clientContext.Site.OpenWeb(parentWeb);

            var lists = hostWeb.Lists;

            _clientContext.Load(lists);
            _clientContext.ExecuteQuery();

            var list = lists.FirstOrDefault(l => l.Title == listName);
            _clientContext.Load(list);

            var defaultView = list.DefaultView;
            foreach (var fieldName in fieldNames)
            {
                defaultView.ViewFields.Add(fieldName);    
            }

            defaultView.Update();

            _clientContext.ExecuteQuery();
        }

        public void RenameField(string listName, string parentWeb,  string originalFieldName, string newFieldName)
        {
            _logger.Information("Renaming field '{1}' to '{2}' in List '{0}'", listName, originalFieldName, newFieldName);

            var hostWeb = string.IsNullOrWhiteSpace(parentWeb)
                ? _clientContext.Site.RootWeb
                : _clientContext.Site.OpenWeb(parentWeb);

            var lists = hostWeb.Lists;
            _clientContext.Load(lists);
            _clientContext.ExecuteQuery();

            var list = lists.FirstOrDefault(l => l.Title == listName);
            _clientContext.Load(list);

            var fields = list.Fields;
            _clientContext.Load(fields);
            _clientContext.ExecuteQuery();

            var field = fields.GetByInternalNameOrTitle(originalFieldName);
            field.Title = newFieldName;
            field.Update();
            _clientContext.Load(field);
            _clientContext.ExecuteQuery();
        }
    }
}
