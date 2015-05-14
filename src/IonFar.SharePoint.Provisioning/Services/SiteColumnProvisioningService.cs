using System;
using System.Linq;
using System.Text;
using IonFar.SharePoint.Provisioning.Infrastructure;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace IonFar.SharePoint.Provisioning.Services
{
    public class SiteColumnProvisioningService : ISiteColumnProvisioningService
    {
        private readonly ClientContext _clientContext;
        private readonly IProvisionLog _logger;

        public SiteColumnProvisioningService(ClientContext clientContext, IProvisionLog logger)
        {
            _clientContext = clientContext;
            _logger = logger;
        }

        public void CreateChoiceField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, string[] choices, ChoiceFormatType format, string defaultValue = null)
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
        }

        public void CreateCurrencyField(string fieldName, string fieldDisplayName, string description, string fieldGroup, bool isRequired, bool isHidden,
            int numberOfDecimalPlaces = 2)
        {
            _logger.Information("Provisioning currency field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var fieldXml = "<Field Type='Currency' Required='" + isRequired  + "' Description='" + description + "' Decimals='" + numberOfDecimalPlaces + "' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "'></Field>";

            var fields = _clientContext.Web.Fields;
            _clientContext.Load(fields);

            fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.ExecuteQuery();
        }

        public void CreateDateField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, bool isDateOnly, string defaultValue = null)
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
        }

        public void CreateHyperlinkField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired,
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
        }

        public void CreateImageField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden)
        {
            _logger.Information("Provisioning image field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var fields = _clientContext.Web.Fields;

            var fieldXml = "<Field Type='Image' Required='" + isRequired + "' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "'></Field>";

            var createdField = fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.Load(fields);
            _clientContext.Load(createdField);
            _clientContext.ExecuteQuery();
        }

        public void CreateLookupField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, string lookupList, string lookupField, bool allowMultipleValues)
        {
            _logger.Information("Provisioning lookup field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var sourceList = _clientContext.Web.Lists.GetByTitle(lookupList);
            var web = _clientContext.Web;
            _clientContext.Load(sourceList);
            _clientContext.Load(web);

            _clientContext.ExecuteQuery();

            var lookupListId = sourceList.Id;
            
            var fields = _clientContext.Web.Fields;

            var fieldXml = "<Field Type='Lookup' Required='" + isRequired + "' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "' List='{" + lookupListId + "}' ShowField='" + lookupField + "' PrependId='TRUE' Mult='" + allowMultipleValues + "' WebId='" + web.Id + "'></Field>";

            var createdField = fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.Load(fields);
            _clientContext.Load(createdField);
            _clientContext.ExecuteQuery();
        }

        public void CreateManagedMetadataField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, bool allowMultipleValues, Guid termStoreId, Guid termSetId, bool isOpen)
        {
            _logger.Information("Provisioning managed metadata field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var fields = _clientContext.Web.Fields;

            var noteFieldId = Guid.NewGuid();

            var noteFieldXml = "<Field ID='{" + noteFieldId + "}' Type='Note' DisplayName='" + fieldDisplayName + "_0' Name='" + fieldName +
                "_0' Group='" + fieldGroup + "' Hidden='TRUE'></Field>";

            var metadataFieldXml = "<Field DisplayName='" + fieldDisplayName + "' Name='" + fieldName + "' Group='" + fieldGroup+ "' Type='TaxonomyFieldTypeMulti' />";
                
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
        }

        public void CreateNoteField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden)
        {
            _logger.Information("Provisioning note field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var fieldXml = "<Field Type='Note' Required='" + isRequired + "' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "'></Field>";

            var fields = _clientContext.Web.Fields;
            _clientContext.Load(fields);

            fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.ExecuteQuery();
        }

        public void CreateTextField(string fieldName, string fieldDisplayName, string description, string fieldGroup, bool isRequired, bool isHidden)
        {
            _logger.Information("Provisioning text field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var fieldXml = "<Field Type='Text' Required='" + isRequired + "' Description='" + description + "' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "'></Field>";

            var fields = _clientContext.Web.Fields;
            _clientContext.Load(fields);

            fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.ExecuteQuery();
        }
        
        public void CreateUserField(string fieldName, string fieldDisplayName, string fieldGroup, FieldUserSelectionMode userSelectionMode, bool isRequired, bool isHidden)
        {
            _logger.Information("Provisioning user field '{0}' to field group '{1}'", fieldName, fieldGroup);

            var fieldXml = "<Field Type='User' UserSelectionMode='" + userSelectionMode + "' Required='" + isRequired + "' DisplayName='" + fieldDisplayName + "' Name='" + fieldName +
                "' Group='" + fieldGroup + "' Hidden='" + isHidden + "'></Field>";

            var fields = _clientContext.Web.Fields;
            _clientContext.Load(fields);

            fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddToNoContentType);

            _clientContext.ExecuteQuery();
        }

        public void DeleteField(string fieldName)
        {
            _logger.Warning("Deleting field '{0}'", fieldName);

            var fields = _clientContext.Web.Fields;       
            var field = fields.GetByTitle(fieldName);
            field.DeleteObject();

            _clientContext.ExecuteQuery();
        }

        public void DeleteFieldsInGroup(string groupName)
        {
            _logger.Warning("Deleting fields in group '{0}'", groupName);
            
            var fields = _clientContext.Web.Fields;
            _clientContext.Load(fields);

            _clientContext.ExecuteQuery();

            var fieldsToDelete = fields.Where(field => field.Group == groupName).ToArray();
            for (var i = 0; i < fieldsToDelete.Length; i++) {
                _logger.Warning("Deleting field '{0}'", fieldsToDelete[i].Title);
                fieldsToDelete[i].DeleteObject();
            }

            _clientContext.ExecuteQuery();
        }
    }
}
