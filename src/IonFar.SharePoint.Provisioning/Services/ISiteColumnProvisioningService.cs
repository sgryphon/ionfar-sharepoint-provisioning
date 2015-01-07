using System;
using Microsoft.SharePoint.Client;

namespace IonFar.SharePoint.Provisioning.Services
{
    public interface ISiteColumnProvisioningService
    {
        void CreateChoiceField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, string[] choices, ChoiceFormatType format, string defaultValue = null);
        void CreateCurrencyField(string fieldName, string fieldDisplayName, string description, string fieldGroup, bool isRequired, bool isHidden, int numberOfDecimalPlaces);  
        void CreateDateField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, bool isDateOnly, string defaultValue = null);
        void CreateHyperlinkField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden);
        void CreateImageField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden);
        void CreateLookupField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, string lookupList, string lookupField, bool allowMultipleValues);
        void CreateManagedMetadataField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, bool allowMultipleValues, Guid termStoreId, Guid termSetId, bool isOpen);
        void CreateNoteField(string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden);
        void CreateTextField(string fieldName, string fieldDisplayName, string description, string fieldGroup, bool isRequired, bool isHidden);
        void CreateUserField(string fieldName, string fieldDisplayName, string fieldGroup, FieldUserSelectionMode userSelectionMode, bool isRequired, bool isHidden);
        void DeleteField(string fieldName);
        void DeleteFieldsInGroup(string groupName);
    }
}