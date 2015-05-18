using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;

namespace IonFar.SharePoint.Provisioning.Services
{
    public interface IContentTypeProvisioningService
    {

        void AddFieldLinkToContentType(string contentTypeId, string fieldInternalNameOrTitle);

        FieldMultiChoice CreateChoiceField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, string[] choices, ChoiceFormatType format, string defaultValue = null);

        ContentType CreateContentType(string contentTypeName, string contentTypeDescription, string contentTypeGroup, string contentTypeId);

        FieldCurrency CreateCurrencyField(string fieldName, string fieldDisplayName, string description, string fieldGroup, bool isRequired, bool isHidden, int numberOfDecimalPlaces);

        FieldDateTime CreateDateField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, bool isDateOnly, string defaultValue = null);

        FieldUrl CreateHyperlinkField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden);

        FieldUrl CreateImageField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden);

        FieldLookup CreateLookupField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, string lookupListTitle, string lookupFieldInternalName, bool allowMultipleValues);

        TaxonomyField CreateManagedMetadataField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, bool allowMultipleValues, Guid termStoreId, Guid termSetId, bool isOpen);

        FieldMultiLineText CreateNoteField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden);

        FieldText CreateTextField(string fieldName, string fieldDisplayName, string description, string fieldGroup, bool isRequired, bool isHidden);

        FieldUser CreateUserField(string fieldName, string fieldDisplayName, string fieldGroup, FieldUserSelectionMode userSelectionMode, bool isRequired, bool isHidden);

        // FieldCalculated
        // FieldChoice
        // FieldComputed
        // FieldGeolocation
        // FieldGuid
        // FieldNumber
        // FieldRatingScale

        void DeleteContentType(string contentTypeId);

        void DeleteField(string fieldName);

        void DeleteFieldsInGroup(string groupName);
    }
}