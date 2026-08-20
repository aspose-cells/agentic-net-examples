// Title: Confirm that a password‑protected XLSX cannot be opened as Excel 97‑2003 (XLS) with Aspose.Cells for .NET
// Description: This C# example creates a workbook, applies password protection, saves it as XLSX, and demonstrates three verification methods: (1) checking the IsEncrypted flag after save, (2) attempting to load the file as Excel97To2003 format and handling the expected exception, and (3) using FileFormatUtil to detect encryption and validate the password via stream. The sample shows how older Excel versions reject the encrypted OOXML file.
// Keywords: Aspose.Cells encryption detection | password protected XLSX compatibility | LoadFormat.Excel97To2003 exception | FileFormatUtil VerifyPassword | IsEncrypted property .NET | secure Excel file handling
// Common Searches: Can an encrypted .xlsx be opened in Excel 97‑2003? | Aspose.Cells detect encrypted workbook before loading | Verify password of protected Excel file without opening | Load encrypted Excel file as older format error | Check IsEncrypted flag after saving workbook
// Developer Intent: Ensure that a workbook encrypted with a password is unreadable by legacy Excel (XLS) versions and that its protection status can be programmatically confirmed.
// Use Cases: Detect encryption status with FileFormatUtil.DetectFileFormat to avoid unnecessary load attempts. | Attempt to open an encrypted XLSX as Excel97To2003, catch the exception, and inform users of incompatibility. | Validate a supplied password using FileFormatUtil.VerifyPassword on a file stream for both correct and incorrect entries.
// AI Prompts: Write C# code that loads an encrypted .xlsx with Aspose.Cells, forces LoadFormat.Excel97To2003, and captures the resulting exception. | Show how to use FileFormatUtil.VerifyPassword on a file stream to test a correct and an incorrect password for a protected workbook. | Explain how to log encryption detection results and password verification outcomes when processing a batch of Excel files with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // This C# example creates a workbook, applies password protection, saves it as XLSX, and demonstrates three verification methods: (1) checking the IsEncrypted flag after save, (2) attempting to load the file as Excel97To2003 format and handling the expected exception, and (3) using FileFormatUtil to detect encryption and validate the password via stream. The sample shows how older Excel versions reject the encrypted OOXML file.
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and add some data
            // -----------------------------------------------------------------
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // -----------------------------------------------------------------
            // 2. Encrypt the workbook with a password
            // -----------------------------------------------------------------
            wb.Settings.Password = "SecretPwd";
            string encryptedPath = "encrypted.xlsx";
            wb.Save(encryptedPath, SaveFormat.Xlsx);

            // Verify that the workbook reports being encrypted
            Console.WriteLine($"After saving, IsEncrypted: {wb.Settings.IsEncrypted}");

            // -----------------------------------------------------------------
            // 3. Load the encrypted workbook with the correct password
            // -----------------------------------------------------------------
            LoadOptions loadWithPwd = new LoadOptions { Password = "SecretPwd" };
            Workbook loadedWithPwd = new Workbook(encryptedPath, loadWithPwd);
            Console.WriteLine($"Loaded with correct password, IsEncrypted: {loadedWithPwd.Settings.IsEncrypted}");
            Console.WriteLine($"Cell A1 value: {loadedWithPwd.Worksheets[0].Cells["A1"].Value}");

            // -----------------------------------------------------------------
            // 4. Attempt to open the encrypted file as an older Excel format (XLS)
            //    Older versions cannot handle the OOXML encryption, so this should fail.
            // -----------------------------------------------------------------
            try
            {
                // Force the loader to treat the file as Excel 97-2003 format
                LoadOptions oldFormatLoad = new LoadOptions(LoadFormat.Excel97To2003);
                Workbook oldFormatWorkbook = new Workbook(encryptedPath, oldFormatLoad);
                Console.WriteLine("Unexpectedly opened encrypted file with old format.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to open encrypted file with older Excel format (as expected).");
                Console.WriteLine($"Exception message: {ex.Message}");
            }

            // -----------------------------------------------------------------
            // 5. Use FileFormatUtil to detect encryption status without opening the file
            // -----------------------------------------------------------------
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(encryptedPath);
            Console.WriteLine($"FileFormatUtil reports IsEncrypted: {formatInfo.IsEncrypted}");

            // -----------------------------------------------------------------
            // 6. Verify password using FileFormatUtil.VerifyPassword (stream based)
            // -----------------------------------------------------------------
            using (Stream stream = File.OpenRead(encryptedPath))
            {
                bool isPwdValid = FileFormatUtil.VerifyPassword(stream, "SecretPwd");
                Console.WriteLine($"VerifyPassword (correct) returned: {isPwdValid}");
            }

            using (Stream stream = File.OpenRead(encryptedPath))
            {
                bool isPwdValid = FileFormatUtil.VerifyPassword(stream, "WrongPwd");
                Console.WriteLine($"VerifyPassword (incorrect) returned: {isPwdValid}");
            }

            // Cleanup
            wb.Dispose();
            loadedWithPwd.Dispose();
        }
    }
}
