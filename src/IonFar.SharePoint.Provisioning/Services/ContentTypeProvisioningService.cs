﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using IonFar.SharePoint.Provisioning.Infrastructure;
using Microsoft.SharePoint.Client;
using System.Text;
using System.Linq;
using Microsoft.SharePoint.Client.Taxonomy;

namespace IonFar.SharePoint.Provisioning.Services
{
    /// <summary>
    /// Service for managing site content types and site columns.
    /// </summary>
    public class ContentTypeProvisioningService : IContentTypeProvisioningService
    {
        private readonly ClientContext _clientContext;
        private readonly IProvisionLog _logger;
        private readonly string[] _contentTypeIds;
        private readonly FieldCollection _fields;

        /// <summary>
        /// Creates a new content type provisioning service
        /// </summary>
        /// <param name="clientContext">Context to use</param>
        /// <param name="logger">(Optional) logger to use; if not specified defaults to TraceSource</param>
        public ContentTypeProvisioningService(ClientContext clientContext, IProvisionLog logger = null)
        {
            _clientContext = clientContext;
            _logger = logger ?? new TraceProvisionLog();

            // Can't cache content types themselves as we're invalidating the context's content type list everytime we add a new one.
            var contentTypes = _clientContext.Site.RootWeb.ContentTypes;
            _clientContext.Load(contentTypes, cts => cts.Include(ct => ct.Id, ct => ct.StringId, ct => ct.Name));

            // Can cache fields as we're not changing the context's field list
            _fields = _clientContext.Site.RootWeb.Fields;
            _clientContext.Load(_fields, fs => fs.Include(f => f.Id, f => f.InternalName));

            _clientContext.ExecuteQuery();

            _contentTypeIds = contentTypes.Select(ct => ct.StringId).ToArray();
        }

        /// <summary>
        /// Adds a site column field to a site content type (in the context web)
        /// </summary>
        /// <param name="contentTypeId">Content type</param>
        /// <param name="fieldInternalNameOrTitle">Internal name, or title, of the site column</param>
        public void AddFieldLinkToContentType(string contentTypeId, string fieldInternalNameOrTitle)
        {
            _logger.Information("Adding field '{0}' to content type id '{1}'", fieldInternalNameOrTitle, contentTypeId);

            var hostWeb = _clientContext.Site.RootWeb;

            var contentTypes = hostWeb.ContentTypes;
            var contentType = contentTypes.GetById(contentTypeId);

            var field = hostWeb.Fields.GetByInternalNameOrTitle(fieldInternalNameOrTitle);

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

        /// <summary>
        /// Creates a site column of type choice (a subtype of multi choice)
        /// </summary>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <param name="choices">Array of string choices</param>
        /// <param name="format">Format the field is displayed, e.g. drop down list</param>
        /// <param name="defaultValue">(Optional) default value of the field</param>
        /// <returns>The created field</returns>
        public FieldMultiChoice CreateChoiceField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, string[] choices, ChoiceFormatType format, string defaultValue = null)
        {
            _logger.Information("Provisioning choice field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var fields = _clientContext.Web.Fields;

            var choicesXml = new StringBuilder();
            foreach (var choice in choices)
            {
                choicesXml.Append("<CHOICE>" + choice + "</CHOICE>");
            }

            var fieldXml = "<Field Type='Choice' Format='" + format + "' Required='" + isRequired + "' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "'>" +
                (!string.IsNullOrWhiteSpace(defaultValue) ? "<Default>" + defaultValue + "</Default>" : "") +
                "<CHOICES>" + choicesXml + "</CHOICES></Field>";

            var createdField = fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.Load(fields);
            _clientContext.Load(createdField);
            _clientContext.ExecuteQuery();

            return (FieldMultiChoice)createdField;
        }

        /// <summary>
        /// Creats a site content type (in the context web)
        /// </summary>
        /// <param name="contentTypeName">Name of the conten type</param>
        /// <param name="contentTypeDescription">Description of the content type</param>
        /// <param name="contentTypeGroup">Group the site content type should appear in</param>
        /// <param name="contentTypeId">ID of the content type; this also determines the parent and inheritance hierarchy</param>
        /// <returns>The created content type</returns>
        public ContentType CreateContentType(string contentTypeName, string contentTypeDescription, string contentTypeGroup, string contentTypeId)
        {
            _logger.Information("Creating Content Type '{0}' in group '{1}'", contentTypeName, contentTypeGroup);

            var hostWeb = _clientContext.Web;

            var contentTypes = hostWeb.ContentTypes;

            var contentTypeCreationInformation = new ContentTypeCreationInformation
            {
                Name = contentTypeName,
                Description = contentTypeDescription,
                Group = contentTypeGroup,
                Id = contentTypeId
            };

            var createdContentType = contentTypes.Add(contentTypeCreationInformation);

            _clientContext.Load(contentTypes);

            _clientContext.ExecuteQuery();

            return createdContentType;
        }

        /// <summary>
        /// Creates a site column of type currency
        /// </summary>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="description">Description of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <param name="numberOfDecimalPlaces">Number of decimal places in the field</param>
        /// <returns>The created field</returns>
        public FieldCurrency CreateCurrencyField(string fieldName, string fieldDisplayName, string description, string fieldGroup, bool isRequired, bool isHidden,
            int numberOfDecimalPlaces = 2)
        {
            _logger.Information("Provisioning currency field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var fieldXml = "<Field Type='Currency' Required='" + isRequired + "' Description='" + description + "' Decimals='" + numberOfDecimalPlaces + "' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "'></Field>";

            var fields = _clientContext.Web.Fields;
            _clientContext.Load(fields);

            var createdField = fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.ExecuteQuery();

            return (FieldCurrency)createdField;
        }

        /// <summary>
        /// Creates a site column of type DateTime
        /// </summary>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <param name="isDateOnly">true if the field is date only; false for date and time</param>
        /// <param name="defaultValue">default value forumla, e.g. "[Today]"</param>
        /// <returns>The created field</returns>
        public FieldDateTime CreateDateField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, bool isDateOnly, string defaultValue = null)
        {
            _logger.Information("Provisioning date field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var dateOnlyAttribute = isDateOnly ? " Format='DateOnly'" : string.Empty;
            var fields = _clientContext.Web.Fields;

            var fieldXml = "<Field Type='DateTime' Required='" + isRequired + "' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "'" + dateOnlyAttribute + ">" + defaultValue + "</Field>";

            var createdField = fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.Load(fields);
            _clientContext.Load(createdField);
            _clientContext.ExecuteQuery();

            return (FieldDateTime)createdField;
        }

        /// <summary>
        /// Creates a site column of type URL, representing a link
        /// </summary>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <returns>The created field</returns>
        public FieldUrl CreateHyperlinkField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired,
            bool isHidden)
        {
            _logger.Information("Provisioning Hyperlink field '{0}' to field group '{1}'", fieldName, fieldGroup);
            var fields = _clientContext.Web.Fields;

            var fieldXml = "<Field Type='URL' Required='" + isRequired + "' Format='Hyperlink' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "'></Field>";

            var createdField = fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.Load(fields);
            _clientContext.Load(createdField);
            _clientContext.ExecuteQuery();

            return (FieldUrl)createdField;
        }

        /// <summary>
        /// Creates a site column of type URL, representing an image
        /// </summary>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <returns>The created field</returns>
        public FieldUrl CreateImageField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden)
        {
            _logger.Information("Provisioning image field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var fields = _clientContext.Web.Fields;

            var fieldXml = "<Field Type='Image' Required='" + isRequired + "' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "'></Field>";

            var createdField = fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.Load(fields);
            _clientContext.Load(createdField);
            _clientContext.ExecuteQuery();

            return (FieldUrl)createdField;
        }

        /// <summary>
        /// Creates a site column of type Lookup
        /// </summary>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <param name="lookupListTitle">Title of the lookup list (in the context web)</param>
        /// <param name="lookupFieldInternalName">Name of the field in the lookup list to display</param>
        /// <param name="allowMultipleValues">true to allow multiple values</param>
        /// <returns>The created field</returns>
        public FieldLookup CreateLookupField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, string lookupListTitle, string lookupFieldInternalName, bool allowMultipleValues)
        {
            _logger.Information("Provisioning lookup field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var sourceList = _clientContext.Web.Lists.GetByTitle(lookupListTitle);
            var web = _clientContext.Web;
            _clientContext.Load(sourceList);
            _clientContext.Load(web);

            _clientContext.ExecuteQuery();

            var lookupListId = sourceList.Id;

            var fields = _clientContext.Web.Fields;

            var fieldXml = "<Field Type='Lookup' Required='" + isRequired + "' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "' List='{" + lookupListId + "}' ShowField='" + lookupFieldInternalName + "' PrependId='TRUE' Mult='" + allowMultipleValues + "' WebId='" + web.Id + "'></Field>";

            var createdField = fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.Load(fields);
            _clientContext.Load(createdField);
            _clientContext.ExecuteQuery();

            return (FieldLookup)createdField;
        }

        /// <summary>
        /// Creates a site column of type managed metadata, as well as the associated hidden note field
        /// </summary>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <param name="allowMultipleValues">true to allow multiple values</param>
        /// <param name="termStoreId">ID of the term store to get values from</param>
        /// <param name="termSetId">ID of the term set to get values from</param>
        /// <param name="isOpen">true if the term set is open and values can be added</param>
        /// <returns>The created field</returns>
        public TaxonomyField CreateManagedMetadataField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, bool allowMultipleValues, Guid termStoreId, Guid termSetId, bool isOpen)
        {
            _logger.Information("Provisioning managed metadata field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var fields = _clientContext.Web.Fields;

            var noteFieldId = Guid.NewGuid();

            var noteFieldXml = "<Field ID='{" + noteFieldId + "}' Type='Note' DisplayName='" + fieldDisplayName + "_0' Name='" + fieldName +
                "_0' Group='" + fieldGroup + "' Hidden='TRUE'></Field>";

            var metadataFieldXml = "<Field DisplayName='" + fieldDisplayName + "' Name='" + fieldName + "' Group='" + fieldGroup + "' Type='TaxonomyFieldTypeMulti' />";

            var noteField = fields.AddFieldAsXml(noteFieldXml, false, AddFieldOptions.AddToNoContentType);
            var metadataField = fields.AddFieldAsXml(metadataFieldXml, false, AddFieldOptions.AddToNoContentType);
            _clientContext.ExecuteQuery();

            var taxonomyField = _clientContext.CastTo<TaxonomyField>(metadataField);
            taxonomyField.SspId = termStoreId;
            taxonomyField.TermSetId = termSetId;
            taxonomyField.TargetTemplate = String.Empty;
            taxonomyField.AnchorId = Guid.Empty;
            taxonomyField.AllowMultipleValues = allowMultipleValues;
            taxonomyField.Open = isOpen;
            taxonomyField.Update();

            _clientContext.ExecuteQuery();

            return taxonomyField;
        }

        /// <summary>
        /// Creates a site column of type multiple lines of text
        /// </summary>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <returns>The created field</returns>
        public FieldMultiLineText CreateNoteField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden)
        {
            _logger.Information("Provisioning note field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var fieldXml = "<Field Type='Note' Required='" + isRequired + "' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "'></Field>";

            var fields = _clientContext.Web.Fields;
            _clientContext.Load(fields);

            var createdField = fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.ExecuteQuery();

            return (FieldMultiLineText)createdField;
        }

        /// <summary>
        /// Creates a site column of type single line of text
        /// </summary>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="description">Description of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <returns>The created field</returns>
        public FieldText CreateTextField(string fieldName, string fieldDisplayName, string description, string fieldGroup, bool isRequired, bool isHidden)
        {
            _logger.Information("Provisioning text field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var fieldXml = "<Field Type='Text' Required='" + isRequired + "' Description='" + description + "' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "'></Field>";

            var fields = _clientContext.Web.Fields;
            _clientContext.Load(fields);

            var createdField = fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.ExecuteQuery();

            return (FieldText)createdField;
        }

        /// <summary>
        /// Creates a site column of type User
        /// </summary>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <param name="userSelectionMode">Whether to select people only, or people and groups</param>
        /// <returns>The created field</returns>
        public FieldUser CreateUserField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, FieldUserSelectionMode userSelectionMode)
        {
            _logger.Information("Provisioning user field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var fieldXml = "<Field Type='User' UserSelectionMode='" + userSelectionMode + "' Required='" + isRequired + "' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "'></Field>";

            var fields = _clientContext.Web.Fields;
            _clientContext.Load(fields);

            var createdField = fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.ExecuteQuery();

            return (FieldUser)createdField;
        }

        /// <summary>
        /// Deletes the specified content type
        /// </summary>
        /// <param name="contentTypeId">ID of the content type</param>
        public void DeleteContentType(string contentTypeId)
        {
            _logger.Warning("Deleting ContentType '{0}'", contentTypeId);

            var hostWeb = _clientContext.Site.RootWeb;

            var contentTypes = hostWeb.ContentTypes;
            var contentTypeToDelete = contentTypes.GetById(contentTypeId);

            contentTypeToDelete.DeleteObject();

            _clientContext.ExecuteQuery();
        }

        /// <summary>
        /// Deletes the specified site column
        /// </summary>
        /// <param name="fieldInternalNameOrTitle">Internal name, or title, of the site column</param>
        public void DeleteField(string fieldInternalNameOrTitle)
        {
            _logger.Warning("Deleting field '{0}'", fieldInternalNameOrTitle);

            var fields = _clientContext.Web.Fields;
            var field = fields.GetByInternalNameOrTitle(fieldInternalNameOrTitle);
            field.DeleteObject();

            _clientContext.ExecuteQuery();
        }

        /// <summary>
        /// Deletes all site columns in the specified group
        /// </summary>
        /// <param name="groupName">Name of the group to delete</param>
        public void DeleteFieldsInGroup(string groupName)
        {
            _logger.Warning("Deleting fields in group '{0}'", groupName);

            var fields = _clientContext.Web.Fields;
            _clientContext.Load(fields);

            _clientContext.ExecuteQuery();

            var fieldsToDelete = fields.Where(field => field.Group == groupName).ToArray();
            for (var i = 0; i < fieldsToDelete.Length; i++)
            {
                _logger.Warning("Deleting field '{0}'", fieldsToDelete[i].Title);
                fieldsToDelete[i].DeleteObject();
            }

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