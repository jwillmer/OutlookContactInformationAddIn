using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace ContactInformations
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        #region Properties

        private Office.IRibbonUI ribbon;
        private ActiveDirectoryHelper activeDirectory;
        private List<ContactInformationProperty> ContactInformationPropertyList;
        private IEnumerable<string> CountryCodeNames;
        private ResourceManager CountryCodeManager;

        #endregion

        #region Constructor

        public Ribbon()
        {
            CountryCodeManager = new ResourceManager("ContactInformations.Properties.CountryCodes", Assembly.GetExecutingAssembly());
            CountryCodeNames = GetResourceNames(CountryCodeManager).OrderBy(x => x);
           
            ContactInformationPropertyList = new List<ContactInformationProperty>();

            ContactInformationPropertyList.Add(new ContactInformationProperty("edbDisplayName", "displayName", "Name"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("edbFirstName", "givenName", "Vorname"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("edbLastname", "sn", "Nachname"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("edbPersonalTitle", "personalTitle", "Anrede"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("edbTitle", "title", "Titel"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("edbOfficeName", "physicalDeliveryOfficeName", "Büroname"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("edbTelephon", "telephoneNumber", "Telefonnummer"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("edbFax", "facsimileTelephoneNumber", "Faxnummer"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("drpdCountry", "c", "Länderabkürzung"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("edbState", "st", "Bundesland"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("edbLocation", "l", "Stadt"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("edbPlz", "postalCode", "Postleitzahl"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("edbStreet", "streetAddress", "Straße"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("edbMail", "mail", "E-Mail-Adresse"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("edbManager", "manager", "Vorgesetzter"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("edbMobile", "mobile", "Handynummer"));
            ContactInformationPropertyList.Add(new ContactInformationProperty("btnContactPic", "thumbnailPhoto", "Profilbild"));

            var adArray = (from property in ContactInformationPropertyList
                           where !property.ActiveDirectoryProperty.Equals(string.Empty)
                           select property.ActiveDirectoryProperty).ToArray();

            activeDirectory = new ActiveDirectoryHelper(adArray);
        }

        #endregion

        #region Delegate calls from UI

        #region Contact-, UpdateClick & DropDownSelect

        public void btnContactClick(Office.IRibbonControl control)
        {
            OpenFileDialog dial = new OpenFileDialog();
            if (dial.ShowDialog() == DialogResult.OK)
            {
                ContactInformationPropertyList.First(property => property.ActiveDirectoryProperty.Equals("thumbnailPhoto"))
                    .NewValue = dial.FileName;
            }
        }

        public void btnUpdateClick(Office.IRibbonControl control)
        {
            if (!activeDirectory.CanConnectToActiveDirectory)
            {
                MessageBox.Show("Es konnte keine Verbindung mit dem Server aufgebaut werden!",
                                "Ferbindungsfehler",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                return;
            }

            ContactInformationPropertyList.ForEach(property =>
            {
                if(!property.NewValue.Equals(string.Empty))
                    if (property.ActiveDirectoryProperty.Equals("thumbnailPhoto"))
                    {
                        try
                        {
                            var fileStream = new FileStream(property.NewValue, FileMode.Open);
                            var binaryReader = new BinaryReader(fileStream);
                            binaryReader.BaseStream.Seek(0, SeekOrigin.Begin);
                            byte[] byteArray = new byte[binaryReader.BaseStream.Length];
                            byteArray = binaryReader.ReadBytes((int)binaryReader.BaseStream.Length);

                            if (!activeDirectory.SetProperty(property.ActiveDirectoryProperty, byteArray))
                                MessageBox.Show("Achten Sie darauf das Ihr Profilbild nicht größer als 100KB ist.",
                                                "Bild konnte nicht aktualisiert werden!",
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Achten Sie darauf das Ihr Profilbild nicht größer als 100KB ist.",
                                                "Bild konnte nicht aktualisiert werden!",
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                        }
                        property.NewValue = string.Empty;
                    }
                    else if (!activeDirectory.SetProperty(property.ActiveDirectoryProperty, property.NewValue))
                        MessageBox.Show(string.Format("Das Textfeld: \"{0}\" konnte nicht aktualisiert werden.", property.DisplayProperty),
                                        "Text konnte nicht aktualisiert werden!",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);
                    else property.NewValue = string.Empty;
            });

            MessageBox.Show("Die Aktualisierung wurde abgeschlossen.",
                            "Profil Aktualisiert",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
        }

        public void drpdSelectedItem(Office.IRibbonControl control, string itemID, int itemIndex)
        {
            OnChange(control, itemID);
        }

        #endregion

        public Bitmap GetButtonImage(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "btnContactPic":
                    var binaryData = activeDirectory.GetValue("thumbnailPhoto") as byte[];
                    if (binaryData != null)
                    {
                        try
                        {
                            var fs = new MemoryStream();
                            var wr = new BinaryWriter(fs);
                            byte[] bb = (byte[])activeDirectory.GetValue("thumbnailPhoto");
                            wr.Write(bb);
                            return new Bitmap(fs);
                        }
                        catch (Exception)
                        {

                        }
                    }
                    return Properties.Resources.defaultContact;
                case "btnUpdate":
                    return Properties.Resources.update;
                default:
                    return null;
            }
        }

        public void OnChange(Office.IRibbonControl control, string text)
        {
            ContactInformationPropertyList.ForEach(property =>
            {
                if (property.ControlProperty.Equals(control.Id))
                    property.NewValue = text;
            });
        }

        public string GetValue(Office.IRibbonControl control)
        {
            string value = string.Empty;
            ContactInformationPropertyList.ForEach(property =>
            {
                if (property.ControlProperty.Equals(control.Id))
                    value = activeDirectory.GetValue(property.ActiveDirectoryProperty).ToString();
            });
            return value;
        }

        public bool GetEnable(Office.IRibbonControl control)
        {
            bool value = false;
            ContactInformationPropertyList.ForEach(property =>
            {
                if (property.ControlProperty.Equals(control.Id))
                    value = activeDirectory.CheckPropertyWritePermission(property.ActiveDirectoryProperty);
            });
            return value;
        }

        #region DropDown

        public string GetItemLabel(Office.IRibbonControl control, int itemIndex)
        {
            return string.Format("{0} - {1}",
                CountryCodeNames.ElementAt(itemIndex), 
                CountryCodeManager.GetString(CountryCodeNames.ElementAt(itemIndex)));
        }

        public int GetItemCount(Office.IRibbonControl control)
        {
            return CountryCodeNames.Count();
        }

        public string GetItemID(Office.IRibbonControl control, int itemIndex)
        {
            return CountryCodeNames.ElementAt(itemIndex);
        }

        public int GetSelectedItemIndex(Office.IRibbonControl control)
        {
            string value = string.Empty;
            ContactInformationPropertyList.ForEach(property =>
            {
                if (property.ControlProperty.Equals(control.Id))
                    value = activeDirectory.GetValue(property.ActiveDirectoryProperty).ToString();
            });
            int position = 0;
            foreach (var element in CountryCodeNames)
            {
                if (element.Equals(value))
                    return position;
                position++;
            }
            return -1;
        }

        #endregion

        #endregion

        #region IRibbonExtensibility-Member

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ContactInformations.Ribbon.xml");
        }

        #endregion

        #region Menübandrückrufe
        //Erstellen Sie hier Rückrufmethoden. Weitere Informationen über das Hinzufügen von Rückrufmethoden erhalten Sie, indem Sie das Menüband-XML-Element im Projektmappen-Explorer markieren und dann F1 drücken.

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helfer

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        private IEnumerable<string> GetResourceNames(ResourceManager resoucreManager)
        {
            ResourceSet rs = resoucreManager.GetResourceSet(CultureInfo.CurrentUICulture, true, true);

            IDictionaryEnumerator ide = rs.GetEnumerator();

            while (ide.MoveNext())
            {
                yield return (string)ide.Key;
            }
        }

        #endregion

        
    }
}
