// Title: Check IsEncrypted Flag After Saving a Password‑Protected Excel Workbook with Aspose.Cells for .NET
// Description: C# example that creates a workbook, sets a password via Workbook.Settings.Password, saves it as an encrypted file, and then verifies the Settings.IsEncrypted property on both the original instance and a reloaded workbook using LoadOptions.
// Keywords: Aspose.Cells | .NET | C# | IsEncrypted | password protection | encrypt Excel workbook | verify encryption flag | Workbook.Settings.Password | LoadOptions password | Excel security
// Common Searches: Aspose.Cells check if workbook is encrypted after save | C# verify Excel password protection with Aspose.Cells | IsEncrypted property Aspose.Cells .NET | How to confirm encrypted Excel file using Aspose.Cells | Load password protected workbook and check encryption flag
// Developer Intent: Ensure that the IsEncrypted property returns true after a workbook is saved with a password and when it is reloaded with the correct password.
// Use Cases: Automated validation that generated reports are correctly password‑protected before distribution. | Unit testing of workbook encryption logic in continuous integration pipelines. | Batch processing scripts that log encryption status for compliance auditing.
// AI Prompts: Generate C# code using Aspose.Cells that creates a workbook, applies a password, saves it, reloads it with LoadOptions, and asserts Settings.IsEncrypted is true. | Write an MSTest unit test that verifies the IsEncrypted flag for a password‑protected workbook saved with Aspose.Cells. | Provide a reusable C# method that accepts a file path and password, loads the workbook with Aspose.Cells, and returns the encryption status.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordVerification
{
    // C# example that creates a workbook, sets a password via Workbook.Settings.Password, saves it as an encrypted file, and then verifies the Settings.IsEncrypted property on both the original instance and a reloaded workbook using LoadOptions.
    class Program
    {
        static void Main()
        {
            // Path for the encrypted workbook
            string encryptedFilePath = "encrypted.xlsx";

            // Create a new workbook and add some data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Password protected workbook");

            // Set the encryption password
            wb.Settings.Password = "Secret123";

            // Save the workbook (it will be encrypted because a password is set)
            wb.Save(encryptedFilePath);

            // Verify the IsEncrypted flag on the original workbook instance
            Console.WriteLine($"After saving, workbook.IsEncrypted: {wb.Settings.IsEncrypted}");

            // Load the encrypted workbook using the correct password
            LoadOptions loadOptions = new LoadOptions { Password = "Secret123" };
            Workbook loadedWb = new Workbook(encryptedFilePath, loadOptions);

            // Verify the IsEncrypted flag on the loaded workbook
            Console.WriteLine($"Loaded workbook IsEncrypted: {loadedWb.Settings.IsEncrypted}");
        }
    }
}
