using CrocoOfficeFileGuard.Core;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using Microsoft.Win32;
using System.IO;
using Xunit;

namespace CrocoOfficeFileGuard.Core.Tests
{
    public class ReservedFileManagerTests
    {
        // ─── HELPER ──────────────────────────────────────────────────────

        /// <summary>
        /// Crea un file .docx temporaneo vuoto
        /// </summary>
        private string CreateTempDocx()
        {
            string path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
            using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                doc.AddMainDocumentPart().Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
            }
            return path;
        }

        /// <summary>
        /// Crea un file .docx temporaneo con la proprietà 'reserved' già impostata (da usare dopo primi test)
        /// </summary>
        private string CreateTempDocxReserved()
        {
            string path = CreateTempDocx();
            ReservedFileManager.ToggleReserved(path); // imposta reserved
            return path;
        }

        // ─── IsSupportedExtension ─────────────────────────────────────────

        [Fact]
        public void IsSupportedExtension_Docx_ReturnsTrue()
            => Assert.True(ReservedFileManager.IsSupportedExtension("file.docx"));

        [Fact]
        public void IsSupportedExtension_Xlsx_ReturnsTrue()
            => Assert.True(ReservedFileManager.IsSupportedExtension("file.xlsx"));

        [Fact]
        public void IsSupportedExtension_Pptx_ReturnsTrue()
            => Assert.True(ReservedFileManager.IsSupportedExtension("file.pptx"));

        [Fact]
        public void IsSupportedExtension_Pdf_ReturnsFalse()
            => Assert.False(ReservedFileManager.IsSupportedExtension("file.pdf"));

        [Fact]
        public void IsSupportedExtension_Txt_ReturnsFalse()
            => Assert.False(ReservedFileManager.IsSupportedExtension("file.txt"));

        [Fact]
        public void IsSupportedExtension_Empty_ReturnsFalse()
            => Assert.False(ReservedFileManager.IsSupportedExtension(""));

        [Fact]
        public void IsSupportedExtension_CaseInsensitive_ReturnsTrue()
            => Assert.True(ReservedFileManager.IsSupportedExtension("file.DOCX"));

        // ─── IsReservedInFile ─────────────────────────────────────────────

        [Fact]
        public void IsReservedInFile_NewFile_ReturnsFalse()
        {
            string path = CreateTempDocx();
            try
            {
                Assert.False(ReservedFileManager.IsReservedInFile(path));
            }
            finally
            {
                File.Delete(path);
                ReservedFileManager.UnregisterReservedFile(path);
            }
        }

        [Fact]
        public void IsReservedInFile_AfterToggle_ReturnsTrue()
        {
            string path = CreateTempDocx();
            try
            {
                ReservedFileManager.ToggleReserved(path);
                Assert.True(ReservedFileManager.IsReservedInFile(path));
            }
            finally
            {
                File.Delete(path);
                ReservedFileManager.UnregisterReservedFile(path);
            }
        }

        [Fact]
        public void IsReservedInFile_AfterDoubleToggle_ReturnsFalse()
        {
            string path = CreateTempDocx();
            try
            {
                ReservedFileManager.ToggleReserved(path); // reserved
                ReservedFileManager.ToggleReserved(path); // unreserved
                Assert.False(ReservedFileManager.IsReservedInFile(path));
            }
            finally
            {
                File.Delete(path);
                ReservedFileManager.UnregisterReservedFile(path);
            }
        }

        [Fact]
        public void IsReservedInFile_NonExistentFile_ReturnsFalse()
            => Assert.False(ReservedFileManager.IsReservedInFile(@"C:\nonexistent\file.docx"));

        // ─── ToggleReserved ───────────────────────────────────────────────

        [Fact]
        public void ToggleReserved_FirstCall_ReturnsTrue()
        {
            string path = CreateTempDocx();
            try
            {
                bool result = ReservedFileManager.ToggleReserved(path);
                Assert.True(result);
            }
            finally
            {
                File.Delete(path);
                ReservedFileManager.UnregisterReservedFile(path);
            }
        }

        [Fact]
        public void ToggleReserved_SecondCall_ReturnsFalse()
        {
            string path = CreateTempDocx();
            try
            {
                ReservedFileManager.ToggleReserved(path);
                bool result = ReservedFileManager.ToggleReserved(path);
                Assert.False(result);
            }
            finally
            {
                File.Delete(path);
                ReservedFileManager.UnregisterReservedFile(path);
            }
        }

        [Fact]
        public void ToggleReserved_UnsupportedExtension_ThrowsNotSupportedException()
        {
            Assert.Throws<NotSupportedException>(() =>
                ReservedFileManager.ToggleReserved("file.pdf"));
        }

        // ─── Registry ─────────────────────────────────────────────────────

        [Fact]
        public void RegisterReservedFile_ThenIsReservedInRegistry_ReturnsTrue()
        {
            string path = Path.Combine(Path.GetTempPath(), "test_reserved.docx");
            try
            {
                ReservedFileManager.RegisterReservedFile(path);
                Assert.True(ReservedFileManager.IsReservedInRegistry(path));
            }
            finally
            {
                ReservedFileManager.UnregisterReservedFile(path);
            }
        }

        [Fact]
        public void UnregisterReservedFile_ThenIsReservedInRegistry_ReturnsFalse()
        {
            string path = Path.Combine(Path.GetTempPath(), "test_reserved.docx");
            ReservedFileManager.RegisterReservedFile(path);
            ReservedFileManager.UnregisterReservedFile(path);
            Assert.False(ReservedFileManager.IsReservedInRegistry(path));
        }

        [Fact]
        public void ToggleReserved_RegistersInRegistry()
        {
            string path = CreateTempDocx();
            try
            {
                ReservedFileManager.ToggleReserved(path);
                Assert.True(ReservedFileManager.IsReservedInRegistry(path));
            }
            finally
            {
                File.Delete(path);
                ReservedFileManager.UnregisterReservedFile(path);
            }
        }

        [Fact]
        public void ToggleReserved_UnregistersFromRegistry()
        {
            string path = CreateTempDocx();
            try
            {
                ReservedFileManager.ToggleReserved(path); // register
                ReservedFileManager.ToggleReserved(path); // unregister
                Assert.False(ReservedFileManager.IsReservedInRegistry(path));
            }
            finally
            {
                File.Delete(path);
                ReservedFileManager.UnregisterReservedFile(path);
            }
        }
    }
}