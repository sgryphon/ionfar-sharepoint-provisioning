using IonFar.SharePoint.Provisioning.Infrastructure;
using Microsoft.SharePoint.Client;

namespace IonFar.SharePoint.Provisioning.Services
{
    public class ContentTypeProvisioningService : IContentTypeProvisioningService
    {
        private readonly ClientContext _clientContext;
        private readonly ILogger _logger;

        public ContentTypeProvisioningService(ClientContext clientContext, ILogger logger)
        {
            _clientContext = clientContext;
            _logger = logger;
        }

        public void CreateContentType(string contentTypeName, string contentTypeDescription, string contentTypeGroup, string contentTypeId)
        {
            _logger.Information("Creating Content Type '{0}' in group '{1}'", contentTypeName, contentTypeGroup);

            var hostWeb = _clientContext.Site.RootWeb;

            var contentTypes = hostWeb.ContentTypes;

            var contentTypeCreationInformation = new ContentTypeCreationInformation
            {
                Name = contentTypeName,
                Description = contentTypeDescription,
                Group = contentTypeGroup,
                Id = contentTypeId
            };

            contentTypes.Add(contentTypeCreationInformation);

            _clientContext.Load(contentTypes);

            _clientContext.ExecuteQuery();
        }

        public void DeleteContentType(string contentTypeId)
        {
            _logger.Warning("Deleting ContentType '{0}'", contentTypeId);

            var hostWeb = _clientContext.Site.RootWeb;

            var contentTypes = hostWeb.ContentTypes;
            var contentTypeToDelete = contentTypes.GetById(contentTypeId);

            contentTypeToDelete.DeleteObject();

            _clientContext.ExecuteQuery();
        }

        public void AddFieldLinkToContentType(string contentTypeId, string fieldName)
        {
            _logger.Information("Adding field '{0}' to content type id '{1}'", fieldName, contentTypeId);

            var hostWeb = _clientContext.Site.RootWeb;

            var contentTypes = hostWeb.ContentTypes;
            var contentType = contentTypes.GetById(contentTypeId);

            var field = hostWeb.Fields.GetByInternalNameOrTitle(fieldName);

            _clientContext.Load(contentType);
            _clientContext.Load(field);

            var fieldLinkCreationInformaton = new FieldLinkCreationInformation
            {
                Field = field,
            };

            contentType.FieldLinks.Add(fieldLinkCreationInformaton);
            contentType.Update(true);

            _clientContext.ExecuteQuery();
        }
    }
}
