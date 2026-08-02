// Title: Validate password protection on an Excel workbook using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, apply an opening password, verify the IsEncrypted flag, confirm that opening the file without a password throws an exception, and successfully load the workbook with LoadOptions.Password to read cell data.
// Keywords: Aspose.Cells password protection | C# verify Excel encryption | IsEncrypted property Aspose | LoadOptions.Password example | exception opening encrypted workbook | .NET Excel file security | validate workbook password
// Common Searches: Aspose.Cells check if workbook is encrypted | C# open password‑protected Excel file with Aspose | how to catch exception when opening encrypted workbook without password | load encrypted Excel file using LoadOptions in .NET | verify Excel file password protection programmatically
// Developer Intent: Confirm that a workbook saved with Settings.Password cannot be opened without the password and can be accessed when the correct password is supplied.
// Use Cases: Create a workbook, set Settings.Password, save, and assert Settings.IsEncrypted is true. | Attempt to instantiate Workbook with only the file path, catch the expected password‑required exception. | Reload the same file using LoadOptions.Password, read a cell value, and verify the encrypted flag remains set.
// AI Prompts: Generate C# code with Aspose.Cells that creates a password‑protected workbook, validates the encryption flag, attempts an unauthorized open to capture the exception, then opens it with the correct password and reads a cell. | Write an MSTest unit test that saves a workbook with Settings.Password, asserts that opening it without LoadOptions throws an Aspose.Cells exception, and confirms successful decryption when the correct password is provided.

using System;
using Aspose.Cells;
using System.IO;

namespace AsposeCellsPasswordValidation
{
    // Demonstrates how to create a workbook, apply an opening password, verify the IsEncrypted flag, confirm that opening the file without a password throws an exception, and successfully load the workbook with LoadOptions.Password to read cell data.
    class Program
    {
        static void Main()
        {
            // Path for the temporary workbook
            string filePath = "encrypted_workbook.xlsx";

            // -------------------------------------------------
            // 1. Create a new workbook and set an opening password
            // -------------------------------------------------
            Workbook wb = new Workbook();                     // create workbook
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data");    // add sample data

            wb.Settings.Password = "Secret123";               // set file encryption password
            wb.Save(filePath);                                // save the encrypted workbook

            // Verify that the workbook is marked as encrypted
            Console.WriteLine($"After saving, IsEncrypted = {wb.Settings.IsEncrypted}");

            // -------------------------------------------------
            // 2. Attempt to open the encrypted workbook without a password
            //    Expect an exception because the password is required
            // -------------------------------------------------
            try
            {
                Workbook withoutPwd = new Workbook(filePath); // try to open without password
                // If no exception, the workbook was opened incorrectly
                Console.WriteLine("ERROR: Workbook opened without password!");
            }
            catch (Exception ex)
            {
                // Expected path: Aspose throws an exception indicating a password is required
                Console.WriteLine($"Opening without password failed as expected: {ex.Message}");
            }

            // -------------------------------------------------
            // 3. Open the workbook with the correct password using LoadOptions
            // -------------------------------------------------
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "Secret123";               // provide the correct password

            Workbook withPwd = new Workbook(filePath, loadOptions); // open with password
            Console.WriteLine($"Opened with password, cell A1 value: {withPwd.Worksheets[0].Cells["A1"].Value}");

            // -------------------------------------------------
            // 4. Confirm that the workbook reports it is encrypted
            // -------------------------------------------------
            Console.WriteLine($"Loaded workbook IsEncrypted = {withPwd.Settings.IsEncrypted}");
        }
    }
}
