using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;

namespace CrocoOfficeFileGuard.Core
{
    public static class ReservedFileManager
    {
        private const string RegistryPath = @"Software\CrocoOfficeFileGuard\ReservedFiles";

        // ─── REGISTRY ────────────────────────────────────────────────────

        public static void RegisterReservedFile(string filePath)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryPath, writable: true))
                key.SetValue(filePath, string.Empty, RegistryValueKind.String);
        }

        public static void UnregisterReservedFile(string filePath)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryPath, writable: true))
                key?.DeleteValue(filePath, throwOnMissingValue: false);
        }

        public static bool IsReservedInRegistry(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) return false;
            try
            {
                string fullPath = Path.GetFullPath(filePath);
                using (var key = Registry.CurrentUser.OpenSubKey(RegistryPath, writable: false))
                    return key?.GetValue(fullPath) != null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
                return false;
            }
        }

        // ─── OFFICE FILE ─────────────────────────────────────────────────

        /// <summary>
        /// Legge la custom property 'reserved' direttamente dal file Office.
        /// Supporta .docx, .xlsx, .pptx
        /// </summary>
        public static bool IsReservedInFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return false;

            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            try
            {
                if (ext == ".docx")
                    return CheckReservedProperty_Word(filePath);
                else if (ext == ".xlsx")
                    return CheckReservedProperty_Excel(filePath);
                else if (ext == ".pptx")
                    return CheckReservedProperty_PowerPoint(filePath);
                else
                    return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
                return false;
            }
        }

        private static bool CheckReservedProperty_Word(string filePath)
        {
            using (var stream = new MemoryStream(File.ReadAllBytes(filePath)))
            using (var doc = WordprocessingDocument.Open(stream, false))
                return HasReservedProperty(doc.CustomFilePropertiesPart);
        }

        private static bool CheckReservedProperty_Excel(string filePath)
        {
            using (var stream = new MemoryStream(File.ReadAllBytes(filePath)))
            using (var doc = SpreadsheetDocument.Open(stream, false))
                return HasReservedProperty(doc.CustomFilePropertiesPart);
        }

        private static bool CheckReservedProperty_PowerPoint(string filePath)
        {
            using (var stream = new MemoryStream(File.ReadAllBytes(filePath)))
            using (var doc = PresentationDocument.Open(stream, false))
                return HasReservedProperty(doc.CustomFilePropertiesPart);
        }

        private static bool HasReservedProperty(CustomFilePropertiesPart customPropsPart)
        {
            if (customPropsPart?.Properties == null) return false;
            return customPropsPart.Properties
                .Elements<CustomDocumentProperty>()
                .Any(p => p.Name?.Value?.Equals("reserved", StringComparison.OrdinalIgnoreCase) == true);
        }

        // ─── TOGGLE ──────────────────────────────────────────────────────

        private static CustomDocumentProperty CreateReservedProperty(int propId)
        {
            var prop = new CustomDocumentProperty
            {
                Name = "reserved",
                FormatId = "{D5CDD505-2E9C-101B-9042-004033C7B5D5}",
                PropertyId = new Int32Value(propId)
            };
            prop.VTBool = new VTBool("true");
            return prop;
        }

        private static int NextPropertyId(Properties customProps)
        {
            int maxId = 1;
            if (customProps != null)
                foreach (var prop in customProps.Elements<CustomDocumentProperty>())
                    if (prop.PropertyId != null && prop.PropertyId.HasValue && prop.PropertyId.Value > maxId)
                        maxId = prop.PropertyId.Value;
            return maxId + 1;
        }

        public static bool ToggleReserved(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".docx" && ext != ".xlsx" && ext != ".pptx")
                throw new NotSupportedException($"Formato non supportato: {ext}");

            if (ext == ".docx") return ToggleReserved_Word(filePath);
            if (ext == ".xlsx") return ToggleReserved_Excel(filePath);
            if (ext == ".pptx") return ToggleReserved_PowerPoint(filePath);
            return false;
        }

        private static bool ToggleReserved_Word(string filePath)
        {
            using (var doc = WordprocessingDocument.Open(filePath, true))
                return ToggleReservedProperty(doc.CustomFilePropertiesPart,
                    () => doc.AddCustomFilePropertiesPart(), filePath);
        }

        private static bool ToggleReserved_Excel(string filePath)
        {
            using (var doc = SpreadsheetDocument.Open(filePath, true))
                return ToggleReservedProperty(doc.CustomFilePropertiesPart,
                    () => doc.AddCustomFilePropertiesPart(), filePath);
        }

        private static bool ToggleReserved_PowerPoint(string filePath)
        {
            using (var doc = PresentationDocument.Open(filePath, true))
                return ToggleReservedProperty(doc.CustomFilePropertiesPart,
                    () => doc.AddCustomFilePropertiesPart(), filePath);
        }

        private static bool ToggleReservedProperty(CustomFilePropertiesPart customPropsPart,
            Func<CustomFilePropertiesPart> addPart, string filePath)
        {
            if (customPropsPart == null)
            {
                customPropsPart = addPart();
                customPropsPart.Properties = new Properties();
                customPropsPart.Properties.Append(CreateReservedProperty(2));
                customPropsPart.Properties.Save();
                RegisterReservedFile(filePath);
                return true; // ora è reserved
            }

            if (customPropsPart.Properties == null)
                customPropsPart.Properties = new Properties();

            var reservedProp = customPropsPart.Properties
                .Elements<CustomDocumentProperty>()
                .FirstOrDefault(p => p.Name?.Value?.Equals("reserved", StringComparison.OrdinalIgnoreCase) == true);

            if (reservedProp != null)
            {
                reservedProp.Remove();
                UnregisterReservedFile(filePath);
                customPropsPart.Properties.Save();
                return false; // ora NON è reserved
            }
            else
            {
                customPropsPart.Properties.Append(CreateReservedProperty(NextPropertyId(customPropsPart.Properties)));
                customPropsPart.Properties.Save();
                RegisterReservedFile(filePath);
                return true; // ora è reserved
            }
        }

        // ─── UTILITY ─────────────────────────────────────────────────────

        public static bool IsSupportedExtension(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) return false;
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            return ext == ".docx" || ext == ".xlsx" || ext == ".pptx";
        }
    }
}