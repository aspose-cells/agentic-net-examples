// Title: Validate that an Aspose.Cells workbook encrypted with a password cannot be opened without it
// Description: C# sample that creates a workbook, encrypts it via Workbook.Settings.Password, saves the file, checks the IsEncrypted flag, attempts to load the file without a password (expecting an exception), then opens it with LoadOptions.Password, reads a cell value and confirms the encryption flag remains true.
// Keywords: Aspose.Cells password protection | encrypted workbook validation | IsEncrypted property | LoadOptions.Password C# | open Excel file without password exception | Aspose.Cells .NET encryption example
// Common Searches: Aspose.Cells verify password protected workbook | how to catch exception when opening encrypted Excel with Aspose.Cells | check IsEncrypted after loading workbook | C# load password protected Excel using Aspose.Cells | test workbook encryption Aspose.Cells .NET
// Developer Intent: Confirm that a workbook saved with a password is inaccessible without the password and can be accessed when the correct password is supplied.
// Use Cases: Create and encrypt a workbook, then assert Settings.IsEncrypted is true. | Attempt to instantiate Workbook from the encrypted file without a password and verify an exception is thrown. | Load the same file with LoadOptions.Password, read cell data, and ensure Settings.IsEncrypted stays true.
// AI Prompts: Generate a C# unit test using Aspose.Cells that asserts opening an encrypted workbook without a password throws an exception. | Provide code to log the IsEncrypted flag before and after loading a password‑protected workbook with Aspose.Cells. | Explain step‑by‑step how to open a password‑protected Excel file in .NET using LoadOptions.Password and handle errors.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordValidation
{
    // C# sample that creates a workbook, encrypts it via Workbook.Settings.Password, saves the file, checks the IsEncrypted flag, attempts to load the file without a password (expecting an exception), then opens it with LoadOptions.Password, reads a cell value and confirms the encryption flag remains true.
    class Program
    {
        static void Main()
        {
            // Path for the temporary workbook
            string filePath = "encrypted_workbook.xlsx";

            // -------------------- Create and encrypt workbook --------------------
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add some data to the first worksheet
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // Set a password to encrypt the workbook
            wb.Settings.Password = "SecretPassword";

            // Save the encrypted workbook
            wb.Save(filePath);

            // Verify that the workbook reports being encrypted
            Console.WriteLine($"IsEncrypted after save: {wb.Settings.IsEncrypted}");

            // -------------------- Attempt to open without password --------------------
            try
            {
                // This should fail because no password is supplied
                Workbook wbNoPassword = new Workbook(filePath);
                Console.WriteLine("Unexpectedly opened workbook without password.");
            }
            catch (Exception ex)
            {
                // Expected failure
                Console.WriteLine($"Failed to open without password: {ex.Message}");
            }

            // -------------------- Open with correct password --------------------
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "SecretPassword";

            // Load the workbook using the correct password
            Workbook wbWithPassword = new Workbook(filePath, loadOptions);
            Console.WriteLine($"Opened with password. Cell A1 value: {wbWithPassword.Worksheets[0].Cells["A1"].StringValue}");

            // Verify that the loaded workbook also reports being encrypted
            Console.WriteLine($"IsEncrypted after load with password: {wbWithPassword.Settings.IsEncrypted}");
        }
    }
}
