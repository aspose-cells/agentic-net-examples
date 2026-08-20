// Title: Verify password‑protected Excel workbook cannot be opened without password using Aspose.Cells for .NET
// Description: This example creates a workbook, writes data, encrypts it via Workbook.Settings.Password, saves it as XLSX, then attempts to load the file without a password (expecting an exception). It reloads the file with LoadOptions.Password, reads the cell value to prove decryption, and checks Workbook.Settings.IsEncrypted to confirm the file remains encrypted.
// Keywords: Aspose.Cells | C# password protection | encrypted workbook | LoadOptions.Password | Workbook.Settings.IsEncrypted | exception opening protected Excel | validate Excel encryption
// Common Searches: Aspose.Cells open encrypted workbook without password | C# check if Excel file is password protected | How to catch exception for protected Excel file in .NET | Verify IsEncrypted flag after loading workbook | Unit test password protection Aspose.Cells
// Developer Intent: Confirm that a workbook encrypted with a password throws an error when opened without providing that password.
// Use Cases: Attempt to load a password‑protected .xlsx file without credentials and handle the expected exception. | Load the same file with LoadOptions.Password, read a cell to ensure successful decryption. | Read Workbook.Settings.IsEncrypted after a successful load to verify the workbook is still marked as encrypted.
// AI Prompts: Generate C# code using Aspose.Cells that creates a workbook, applies a password, saves it, and demonstrates that opening it without the password raises an exception. | Write a C# unit test with Aspose.Cells that asserts an exception is thrown when loading a password‑protected workbook without providing a password. | Explain how to use Workbook.Settings.IsEncrypted to confirm a workbook remains encrypted after loading it with the correct password.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordValidation
{
    // This example creates a workbook, writes data, encrypts it via Workbook.Settings.Password, saves it as XLSX, then attempts to load the file without a password (expecting an exception). It reloads the file with LoadOptions.Password, reads the cell value to prove decryption, and checks Workbook.Settings.IsEncrypted to confirm the file remains encrypted.
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

            // -----------------------------------------------------------------
            // 3. Save the encrypted workbook
            // -----------------------------------------------------------------
            string encryptedPath = "encrypted.xlsx";
            wb.Save(encryptedPath, SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // 4. Attempt to open the encrypted workbook WITHOUT providing a password
            //    Expect an exception because the file is protected.
            // -----------------------------------------------------------------
            try
            {
                // This load does NOT supply a password, so it should fail.
                Workbook wbWithoutPwd = new Workbook(encryptedPath);
                Console.WriteLine("ERROR: Workbook opened without password (unexpected).");
            }
            catch (Exception ex)
            {
                // Expected path: an exception is thrown indicating the file is encrypted.
                Console.WriteLine("Expected exception when opening without password: " + ex.Message);
            }

            // -----------------------------------------------------------------
            // 5. Open the encrypted workbook WITH the correct password
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "SecretPwd"
            };
            Workbook wbWithPwd = new Workbook(encryptedPath, loadOptions);

            // Verify that the workbook is indeed decrypted and data is accessible
            string cellValue = wbWithPwd.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine("Cell A1 value after providing password: " + cellValue);

            // -----------------------------------------------------------------
            // 6. Additional check: confirm the workbook reports it is encrypted
            // -----------------------------------------------------------------
            Console.WriteLine("IsEncrypted flag after loading with password: " + wbWithPwd.Settings.IsEncrypted);
        }
    }
}
