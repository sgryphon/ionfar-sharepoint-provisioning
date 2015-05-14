using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using IonFar.SharePoint.Provisioning.Infrastructure;
using Microsoft.SharePoint.Client;

namespace IonFar.SharePoint.Provisioning.Services
{
    public class ContentTypeProvisioningService : IContentTypeProvisioningService
    {
        private readonly ClientContext _clientContext;
        private readonly IProvisionLog _logger;
        private readonly string[] _contentTypeIds;
        private readonly FieldCollection _fields;

        public ContentTypeProvisioningService(ClientContext clientContext, IProvisionLog logger)
        {
            _clientContext = clientContext;
            _logger = logger;

            // Can't cache content types themselves as we're invalidating the context's content type list everytime we add a new one.
            var contentTypes = _clientContext.Site.RootWeb.ContentTypes;
            _clientContext.Load(contentTypes, cts => cts.Include(ct => ct.Id, ct => ct.StringId, ct => ct.Name));

            // Can cache fields as we're not changing the context's field list
            _fields = _clientContext.Site.RootWeb.Fields;
            _clientContext.Load(_fields, fs => fs.Include(f => f.Id, f => f.InternalName));

            _clientContext.ExecuteQuery();

            _contentTypeIds = contentTypes.Select(ct => ct.StringId).ToArray();
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

        public void Provision(IEnumerable<ContentTypeDescriptor> contentTypeDescriptor)
        {
            _fields.RefreshLoad();
            _clientContext.ExecuteQuery();
            contentTypeDescriptor.ToList().ForEach(Provision);
        }

        public void Provision(ContentTypeDescriptor contentTypeDescriptor)
        {
            _logger.Information("Creating Content Type '{0}' in group '{1}'", contentTypeDescriptor.Name, contentTypeDescriptor.Group);

            if (_contentTypeIds.Any(id => id == contentTypeDescriptor.Id))
            {
                _logger.Warning("  Content Type '{0}' in group '{1}' already exists. Skipping.", contentTypeDescriptor.Name, contentTypeDescriptor.Group);
                return;
            }

            var contentTypes = _clientContext.Site.RootWeb.ContentTypes;

            var contentType = contentTypes.Add(new ContentTypeCreationInformation
            {
                Name = contentTypeDescriptor.Name,
                Description = contentTypeDescriptor.Description,
                Group = contentTypeDescriptor.Group,
                Id = contentTypeDescriptor.Id
            });
            
            foreach (var fieldReference in contentTypeDescriptor.Fields)
            {
                var fieldLinkCreationInformaton = _fields
                    .Where(f => f.InternalName == fieldReference.Name)
                    .Select(f => new FieldLinkCreationInformation { Field = f })
                    .FirstOrDefault();


                if (fieldLinkCreationInformaton == null)
                {
                    _clientContext.Load(_fields);
                    _clientContext.ExecuteQuery();
                    fieldLinkCreationInformaton = _fields
                        .Where(f => f.InternalName == fieldReference.Name)
                        .Select(f => new FieldLinkCreationInformation { Field = f })
                        .FirstOrDefault();
                }

                if (fieldLinkCreationInformaton == null)
                {
                    throw new Exception(string.Format(
                        CultureInfo.InvariantCulture, 
                        "Field '{0}' cannot be added to Content Type '{1}'.  Field does not exist",
                        fieldReference.Name,
                        contentType.Name));
                }

                _logger.Information("Adding field '{0}' to content type.", fieldReference.Name);

                var fieldLink = contentType.FieldLinks.Add(fieldLinkCreationInformaton);
                fieldLink.Hidden = fieldReference.Status == ContentTypeFieldStatus.Hidden;
                fieldLink.Required = fieldReference.Status == ContentTypeFieldStatus.Required;

                contentType.Update(true);
            }

            _clientContext.ExecuteQuery();
        }

        //public void Unprovision(ContentTypeDescriptor contentTypeDescriptor)
        //{
        //    var web = _clientContext.Web;
        //    var contentType = web.GetContentTypeByName(contentTypeDescriptor.Name);
        //    if (contentType == null) return;
        //    contentType.DeleteObject();
        //    _clientContext.ExecuteQuery();

        //    /* clear the cached content-type id */
        //    var index = Array.IndexOf(_contentTypeIds, contentTypeDescriptor.Id);
        //    _contentTypeIds[index] = null;
        //}

    }

}