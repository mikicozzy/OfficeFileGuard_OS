using CrocoOfficeFileGuard.Core;
using System;
using System.IO;

namespace CrocoOfficeFileReserveToggle.Cli
{
    internal class Program
    {
        static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Error: Pathfile not specified");
                return -1;
            }

            string filePath;
            try
            {
                filePath = Path.GetFullPath(args[0]);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: Invalid path - {ex.Message}");
                return -1;
            }

            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                Console.WriteLine("Error: File not found or wrong path");
                return -1;
            }

            if (!ReservedFileManager.IsSupportedExtension(filePath))
            {
                Console.WriteLine("Error: File format not supported");
                return -1;
            }

            try
            {
                bool isNowReserved = ReservedFileManager.ToggleReserved(filePath);
                Console.WriteLine(isNowReserved
                    ? $"File marked as RESERVED: {filePath}"
                    : $"File UNRESERVED: {filePath}");
                return isNowReserved ? 1 : 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return -1;
            }
        }
    }
}


/*
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//using SharedInterfaces;  // import SharedInterfaces

namespace CrocoOfficeFileReserveToggle.Cli
{
    internal class Program
    {

        /// <summary>
        /// Valid for Office files only
        /// Return a new CustomDocumentProperty named 'reserved' with value set to "true"
        /// PARAM:
        ///     propId: the property ID
        /// </summary>
        private static CustomDocumentProperty CreateReservedProperty(int propId)
        {
            CustomDocumentProperty newReservedProp = new CustomDocumentProperty
            {
                Name = "reserved",
                FormatId = "{D5CDD505-2E9C-101B-9042-004033C7B5D5}",    // GUID standard per le proprietà custom
                PropertyId = new Int32Value(propId)
            };
            newReservedProp.VTBool = new VTBool("true");    // sets the 'reserved' value to true
            return newReservedProp;
        }


        /// <summary>
        /// Return the next PropertyId (as the max PropertyID + 1)
        /// PropertyID cannot be less than 2
        /// PARAM:
        ///     - customProps: the Properties collection to check
        /// </summary>
        private static int NextPropertyId(Properties customProps)
        {
            int maxId = 1;
            if (customProps != null)
                foreach (var prop in customProps.Elements<CustomDocumentProperty>())
                {
                    if (prop.PropertyId != null && prop.PropertyId.HasValue)
                    {
                        int currentId = prop.PropertyId.Value;  // getting current PropertyId and casting in int
                        if (currentId > maxId)  maxId = currentId;
                    }
                }
            return maxId + 1;
        }


        /// <summary>
        /// Adds the specified file path to the registry 
        /// in the current user's registry under CrocoOfficeFileGuard.
        /// </summary>
        private static void RegisterReservedFile(string filePath)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\CrocoOfficeFileGuard\ReservedFiles", writable: true))
                {
                    key.SetValue(filePath, string.Empty, RegistryValueKind.String);
                }
                Console.WriteLine($"[RegisterReservedFile] Registered: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RegisterReservedFile] Error registering '{filePath}': {ex.Message}");
            }
}


        /// <summary>
        /// Removes the specified file path from the registry 
        /// in the current user's registry under CrocoOfficeFileGuard.
        /// </summary>
        private static void UnregisterReservedFile(string filePath)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\CrocoOfficeFileGuard\ReservedFiles", writable: true))
                {
                    key?.DeleteValue(filePath, throwOnMissingValue: false);
                }
                Console.WriteLine($"[UnregisterReservedFile] Unregistered: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[UnregisterReservedFile] Error unregistering '{filePath}': {ex.Message}");
            }
        }


        /// <summary>
        /// Toggle the existance of the Custom Property (only for office files) 'reserved' (.docx, .xlsx, .pptx)
        /// FOR NOW: ONLY WORD FILE AND THE PROPERTY IS CREATED AND THEN TOGGLED TO TRUE OR FALSE
        /// </summary>
        static int Main(string[] args)
        {
            if (args.Length == 0) { Console.WriteLine("Error: Pathfile not specified"); return -1; }
            string filePath = Path.GetFullPath(args[0]);   // get the full path from args[0] (args[0] could be a full or a relative path)
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))  {Console.WriteLine("Error: File not found or wrong path");  return -1;}
            if (!filePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))  {Console.WriteLine("Error: File format not supported");  return -1;}

            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))  // (using è equivalente di with) opens the .docx file in write mode
                {
                    CustomFilePropertiesPart customPropsPart = wordDoc.CustomFilePropertiesPart;   // get the custom properties part

                    // --- IF THE CUSTOM PROPERTIES PART DOESN'T EXIST (NO CUSTOM) ---
                    if (customPropsPart == null)
                    {
                        customPropsPart = wordDoc.AddCustomFilePropertiesPart();    // adds a custom properties part
                        customPropsPart.Properties = new Properties();  // initializes the custom properties collection
                        customPropsPart.Properties.Append(CreateReservedProperty(2)); // adds the new 'reserved' property to the collection (2 is the first available ID)
                        customPropsPart.Properties.Save();  // saves the changes
                        // and
                        RegisterReservedFile(filePath);  // registers the file in the registry
                        return 1;
                    }

                    // --- THE CUSTOM PROPERTIES PART EXIST ---
                    Properties customProps = customPropsPart.Properties;    // gets the custom properties collection
                    if (customProps == null)    // the collection 'Properties' could be null even if the Part exists
                    {
                        customPropsPart.Properties = new Properties();  // initializes the custom properties collection
                        customProps = customPropsPart.Properties;
                    }
                    // Looking for 'reserved' property (ignoring case)
                    CustomDocumentProperty reservedProp = customPropsPart.Properties.Elements<CustomDocumentProperty>()
                        .FirstOrDefault(p => p.Name.Value.Equals("reserved", StringComparison.OrdinalIgnoreCase));

                    if (reservedProp != null)  // if the 'reserved' property exists
                    {
                        reservedProp.Remove();  // removes the existing 'reserved' property
                        UnregisterReservedFile(filePath);    // unregisters the file from the registry

                        // getting the 'reserved' value
                        //bool isCurrentlyTrue = reservedProp.VTBool != null && reservedProp.VTBool.InnerText.Equals("true", StringComparison.OrdinalIgnoreCase);
                        //if (isCurrentlyTrue) reservedProp.VTBool = new VTBool("false");    // toggles the value to false
                        //else reservedProp.VTBool = new VTBool("true"); // toggles the value to true
                    }
                    else       // the 'reserved' property does not exist (but the custom properties part exists)
                    {
                        int propId = NextPropertyId(customProps);
                        customPropsPart.Properties.Append(CreateReservedProperty(propId)); // adds the new 'reserved' property to the collection
                        RegisterReservedFile(filePath);   // registers the file in the registry
                    }

                    // Salva le modifiche (il 'using' lo farà automaticamente quando chiude wordDoc)
                    customPropsPart.Properties.Save();
                    return 0;
                }
            }
            catch (Exception ex)
            {
                // Gestione dell'eccezione (log, rilancio, ecc.)
                //throw new ApplicationException("An error occurred while toggling the 'reserved' property.", ex);
                Console.WriteLine($"Errore in ToggleReservedProperty: {ex.Message}");
                return -1;
            }
        }
    }
}
*/