using IonFar.SharePoint.Provisioning.Services;
using System.Collections.Generic;

using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;

namespace IonFar.SharePoint.Provisioning.Services
{
    public interface IContentTypeProvisioningService
    {
        void Provision(IEnumerable<ContentTypeDescriptor> contentTypeDescriptor);
        void Provision(ContentTypeDescriptor contentTypeDescriptor);

        /// <summary>
        /// Adds a site column field to a site content type (in the context web)
        /// </summary>
        /// <param name="contentTypeId">Content type</param>
        /// <param name="fieldInternalNameOrTitle">Internal name, or title, of the site column</param>
        void AddFieldLinkToContentType(string contentTypeId, string fieldInternalNameOrTitle);

        /// <summary>
        /// Creates a site column of type boolean
        /// </summary>
        /// <param name="id">Unique ID of the field</param>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <param name="defaultValue">(Optional) default value of the field</param>
        /// <returns>The created field</returns>
        Field CreateBooleanField(Guid id, string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, bool? defaultValue = null);

        /// <summary>
        /// Creates a site column of type choice (a subtype of multi choice)
        /// </summary>
        /// <param name="id">Unique ID of the field</param>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <param name="choices">Array of string choices</param>
        /// <param name="format">Format the field is displayed, e.g. drop down list</param>
        /// <param name="defaultValue">(Optional) default value of the field</param>
        /// <returns>The created field</returns>
        FieldMultiChoice CreateChoiceField(Guid id, string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, string[] choices, ChoiceFormatType format, string defaultValue = null);

        /// <summary>
        /// Creats a site content type (in the context web)
        /// </summary>
        /// <param name="contentTypeId">ID of the content type; this also determines the parent and inheritance hierarchy</param>
        /// <param name="contentTypeName">Name of the conten type</param>
        /// <param name="contentTypeDescription">Description of the content type</param>
        /// <param name="contentTypeGroup">Group the site content type should appear in</param>
        /// <returns>The created content type</returns>
        ContentType CreateContentType(string contentTypeId, string contentTypeName, string contentTypeDescription, string contentTypeGroup);

        /// <summary>
        /// Creates a site column of type currency
        /// </summary>
        /// <param name="id">Unique ID of the field</param>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="description">Description of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <param name="numberOfDecimalPlaces">Number of decimal places in the field</param>
        /// <returns>The created field</returns>
        FieldCurrency CreateCurrencyField(Guid id, string fieldName, string fieldDisplayName, string description, string fieldGroup, bool isRequired, bool isHidden, int numberOfDecimalPlaces);

        /// <summary>
        /// Creates a site column of type DateTime
        /// </summary>
        /// <param name="id">Unique ID of the field</param>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <param name="isDateOnly">true if the field is date only; false for date and time</param>
        /// <param name="defaultValue">default value forumla, e.g. "[Today]"</param>
        /// <returns>The created field</returns>
        FieldDateTime CreateDateField(Guid id, string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, bool isDateOnly, string defaultValue = null);

        /// <summary>
        /// Creates a site column of type URL, representing a link
        /// </summary>
        /// <param name="id">Unique ID of the field</param>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <returns>The created field</returns>
        FieldUrl CreateHyperlinkField(Guid id, string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden);

        /// <summary>
        /// Creates a site column of type URL, representing an image
        /// </summary>
        /// <param name="id">Unique ID of the field</param>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <returns>The created field</returns>
        FieldUrl CreateImageField(Guid id, string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden);

        /// <summary>
        /// Creates a site column of type Lookup
        /// </summary>
        /// <param name="id">Unique ID of the field</param>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <param name="lookupListTitle">Title of the lookup list (in the context web)</param>
        /// <param name="lookupFieldInternalName">Name of the field in the lookup list to display</param>
        /// <param name="allowMultipleValues">true to allow multiple values</param>
        /// <returns>The created field</returns>
        FieldLookup CreateLookupField(Guid id, string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, string lookupListTitle, string lookupFieldInternalName, bool allowMultipleValues);

        /// <summary>
        /// Creates a site column of type managed metadata, as well as the associated hidden note field
        /// </summary>
        /// <param name="id">Unique ID of the field</param>
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
        TaxonomyField CreateManagedMetadataField(Guid id, string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, bool allowMultipleValues, Guid termStoreId, Guid termSetId, bool isOpen);

        /// <summary>
        /// Creates a site column of type multiple lines of text
        /// </summary>
        /// <param name="id">Unique ID of the field</param>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <returns>The created field</returns>
        FieldMultiLineText CreateNoteField(Guid id, string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden);

        /// <summary>
        /// Creates a site column of type single line of text
        /// </summary>
        /// <param name="id">Unique ID of the field</param>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="description">Description of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <param name="decimalPlaces">Number of decimal places to use</param>
        /// <returns>The created field</returns>
        FieldNumber CreateNumberField(Guid id, string fieldName, string fieldDisplayName, string description, string fieldGroup, bool isRequired, bool isHidden, int decimalPlaces);

        /// <summary>
        /// Creates a site column of type single line of text
        /// </summary>
        /// <param name="id">Unique ID of the field</param>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="description">Description of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <returns>The created field</returns>
        FieldText CreateTextField(Guid id, string fieldName, string fieldDisplayName, string description, string fieldGroup, bool isRequired, bool isHidden);

        /// <summary>
        /// Creates a site column of type User
        /// </summary>
        /// <param name="id">Unique ID of the field</param>
        /// <param name="fieldName">Internal name of the site column</param>
        /// <param name="fieldDisplayName">Display name (title) of the site column</param>
        /// <param name="fieldGroup">Group the site column should appear in</param>
        /// <param name="isRequired">true to default the column to be mandatory</param>
        /// <param name="isHidden">true to create a hidden column</param>
        /// <param name="userSelectionMode">Whether to select people only, or people and groups</param>
        /// <returns>The created field</returns>
        FieldUser CreateUserField(Guid id, string fieldName, string fieldDisplayName, string fieldGroup, bool isRequired, bool isHidden, FieldUserSelectionMode userSelectionMode);

        // FieldCalculated
        // FieldChoice
        // FieldComputed
        // FieldGeolocation
        // FieldGuid
        // FieldNumber
        // FieldRatingScale

        /// <summary>
        /// Deletes the specified content type
        /// </summary>
        /// <param name="contentTypeId">ID of the content type</param>
        void DeleteContentType(string contentTypeId);

        /// <summary>
        /// Deletes the specified site column
        /// </summary>
        /// <param name="fieldInternalNameOrTitle">Internal name, or title, of the site column</param>
        void DeleteField(string fieldInternalNameOrTitle);

        /// <summary>
        /// Deletes all site columns in the specified group
        /// </summary>
        /// <param name="groupName">Name of the group to delete</param>
        void DeleteFieldsInGroup(string groupName);
    }
}