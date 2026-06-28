using System;
using System.IO;

namespace CrocoIconOverlayShellExtension
{
    internal static class IconOverlayHelper
    {
        /// <summary>
        /// Central decision point to determine if an overlay must be shown for a path.
        /// Suitable to be used from IShellIconOverlayIdentifier.IsMemberOf or SharpShell overlay hook.
        /// </summary>
        public static bool IsOverlayNeeded(string filePath)
        {
            // fast pre-filter: only supported extensions
            if (!ReservedFileRegistry.IsSupportedExtension(filePath))
                return false;

            // normalized path check against registry
            return ReservedFileRegistry.IsReserved(filePath);
        }
    }
}
