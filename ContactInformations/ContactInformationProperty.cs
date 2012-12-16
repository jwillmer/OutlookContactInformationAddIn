
namespace ContactInformations
{
    class ContactInformationProperty
    {
        public string ActiveDirectoryProperty { get; private set; }
        
        public string ControlProperty { get; private set; }
        
        public string DisplayProperty { get; private set; }

        public string NewValue { get; set; }

        public ContactInformationProperty(string controlProperty, string activeDirectoryProperty, string displayProperty)
        {
            ActiveDirectoryProperty = activeDirectoryProperty;
            ControlProperty = controlProperty;
            DisplayProperty = displayProperty;
            NewValue = string.Empty;
        }
    }
}
