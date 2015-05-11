using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;

namespace IonFar.SharePoint.Provisioning.Services
{
    public abstract class FieldDescriptor 
    {
        protected FieldDescriptor(
            string group, 
            string name, 
            string displayName, 
            bool isRequired, 
            bool isHidden, 
            string description)
        {
            Group = group; 
            Name = name;
            DisplayName = displayName;
            IsRequired = isRequired;
            IsHidden = isHidden;
            Description = description;
        }

        public string Group { get; private set; }
        public string Name { get; private set; }
        public string DisplayName { get; private set; }
        public string Description { get; private set; }
        public bool IsRequired { get; private set; } 
        public bool IsHidden { get; private set; }
    }

    public sealed class DateTimeFieldDescriptor : FieldDescriptor
    {
        public DateTimeFieldDescriptor(
            string group, 
            string name, 
            string displayName, 
            bool isRequired, 
            bool isHidden, 
            string description,
            bool isDateOnly)
            : base(group, name, displayName, isRequired, isHidden, description)
        {
            IsDateOnly = isDateOnly;
        }

        public bool IsDateOnly { get; private set; }
    }

    public sealed class TextFieldDescriptor : FieldDescriptor
    {
        public TextFieldDescriptor(
            string group, 
            string name, 
            string displayName, 
            bool isRequired, 
            bool isHidden, 
            string description)
            : base(group, name, displayName, isRequired, isHidden, description)
        {
        }
    }

    public sealed class NoteFieldDescriptor : FieldDescriptor
    {
        public NoteFieldDescriptor(
            string group, 
            string name, 
            string displayName, 
            bool isRequired, 
            bool isHidden, 
            string description)
            : base(group, name, displayName, isRequired, isHidden, description)
        {
        }
    }

    public sealed class CurrencyFieldDescriptor : FieldDescriptor
    {
        public CurrencyFieldDescriptor(
            string group, 
            string name, 
            string displayName, 
            bool isRequired, 
            bool isHidden, 
            string description,
            int numberOfDecimalPlaces)
            : base(group, name, displayName, isRequired, isHidden, description)
        {
            NumberOfDecimalPlaces = numberOfDecimalPlaces;
        }

        public int NumberOfDecimalPlaces { get; private set; }
    }

    public sealed class ChoiceFieldDescriptor : FieldDescriptor
    {
        public ChoiceFieldDescriptor(
            string group, 
            string name, 
            string displayName, 
            bool isRequired, 
            bool isHidden, 
            string description,
            IEnumerable<string> choices,
            string choiceDefaultValue,
            ChoiceFormatType choiceFormat)
            : base(group, name, displayName, isRequired, isHidden, description)
        {
            Choices = choices;
            ChoiceDefaultValue = choiceDefaultValue;
            ChoiceFormat = choiceFormat;
        }

        public IEnumerable<string> Choices { get; private set; }
        public string ChoiceDefaultValue { get; private set; }
        public ChoiceFormatType ChoiceFormat { get; private set; }
    }

    public sealed class UserFieldDescriptor : FieldDescriptor
    {
        public FieldUserSelectionMode UserSelectionMode { get; private set; }

        public UserFieldDescriptor(
            string group,
            string name,
            string displayName,
            bool isRequired,
            bool isHidden,
            string description,
            FieldUserSelectionMode userSelectionMode)
            : base(group, name, displayName, isRequired, isHidden, description)
        {
            UserSelectionMode = userSelectionMode;
        }
    }

    public sealed class NumberFieldDescriptor : FieldDescriptor
    {
        public NumberFieldDescriptor(
            string group, 
            string name, 
            string displayName, 
            bool isRequired, 
            bool isHidden, 
            string description,
            int numberOfDecimalPlaces,
            bool showAsPercentage)
            : base(group, name, displayName, isRequired, isHidden, description)
        {
            NumberOfDecimalPlaces = numberOfDecimalPlaces;
            ShowAsPercentage = showAsPercentage;
        }

        public int NumberOfDecimalPlaces { get; private set; }
        public bool ShowAsPercentage { get; private set; }
    }

    public sealed class BooleanFieldDescriptor : FieldDescriptor
    {
        public BooleanFieldDescriptor(
            string group, 
            string name, 
            string displayName, 
            bool isRequired, 
            bool isHidden, 
            string description,
            bool defaultValue)
            : base(group, name, displayName, isRequired, isHidden, description)
        {
            DefaultValue = defaultValue;
        }

        public bool DefaultValue { get; set;}
    }

    public sealed class TaxonomyFieldDescriptor : FieldDescriptor
    {
        public TaxonomyFieldDescriptor(
            string group, 
            string name,
            string displayName, 
            bool isRequired, 
            bool isHidden, 
            string description,
            bool allowMultipleValues, 
            Guid termSetId, 
            bool isOpen,
            Guid textFieldId)
            : base(group, name, displayName, isRequired, isHidden, description)
        {
            AllowMultipleValues = allowMultipleValues;
            TermSetId = termSetId;
            IsOpen = isOpen;
            TextFieldId = textFieldId;
        } 

        public bool AllowMultipleValues { get; private set; }
        public Guid TermSetId { get; private set; }
        public bool IsOpen { get; private set; }
        public Guid TextFieldId { get; private set; }
    }

    public sealed class LookupFieldDescriptor : FieldDescriptor
    {
        public LookupFieldDescriptor(
            string group, 
            string name,
            string displayName, 
            bool isRequired, 
            bool isHidden, 
            string description,
            bool allowMultipleValues,
            string lookupList,
            string lookupField)
            : base(group, name, displayName, isRequired, isHidden, description)
        {
            AllowMultipleValues = allowMultipleValues;
            LookupList = lookupList;
            LookupField = lookupField;
        }

        public bool AllowMultipleValues { get; private set; }
        public string LookupList { get; private set; }
        public string LookupField { get; private set; }
    }

    public sealed class ComputedColumnFieldDescriptor : FieldDescriptor
    {
        public ComputedColumnFieldDescriptor(
            string group, 
            string name,
            string displayName, 
            bool isRequired, 
            bool isHidden, 
            string description,
            string resultType,
            string formula,
            IEnumerable<string> fieldReferences)
            : base(group, name, displayName, isRequired, isHidden, description)
        {
            ResultType = resultType;
            Formula = formula;
            FieldReferences = fieldReferences;
        }

        public ComputedColumnFieldDescriptor ChangeBaseUrl(string initial, string final)
        {
            Formula = Formula.Replace(initial, final);
            return this;
        }

        public string ResultType { get; private set; }
        public string Formula { get; private set; }
        public IEnumerable<string> FieldReferences { get; private set; }
    }
}

