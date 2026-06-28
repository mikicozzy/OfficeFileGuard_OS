using Microsoft.Win32;
using System;
using System.IO;

namespace CrocoIconOverlayShellExtension
{
    internal static class ReservedFileRegistry
    {
        private const string RegistryPath = @"Software\CrocoOfficeFileGuard\ReservedFiles";


        /// <summary>
        /// Returns true if the filepath is present in the ReservedFiles of the registry, false if not.
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool IsReserved(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return false;
            try
            {
                string fullPath = Path.GetFullPath(filePath);
                using (var key = Registry.CurrentUser.OpenSubKey(RegistryPath, writable: false))
                {
                    if (key == null)
                        return false;
                    return key.GetValue(fullPath) != null;
                }
            }
            catch
            {
                return false;
            }
        }


        /// <summary>
        /// Returns True id the filePath has docx, xlsx, pptx extension; false if not
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool IsSupportedExtension(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return false;
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            return ext == ".docx" || ext == ".xlsx" || ext == ".pptx";
        }


        /// <summary>
        /// Central decision point to determine if an overlay must be shown for a path.
        /// Suitable to be used from IShellIconOverlayIdentifier.IsMemberOf or SharpShell overlay hook.
        /// </summary>
        public static bool IsOverlayNeeded(string filePath)
        {
            // fast pre-filter: only supported extensions
            if (!ReservedFileRegistry.IsSupportedExtension(filePath))
                return false;

            return ReservedFileRegistry.IsReserved(filePath);
        }
    }

}