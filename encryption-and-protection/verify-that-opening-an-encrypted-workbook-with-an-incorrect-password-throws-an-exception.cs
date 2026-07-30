// Title: C# – Throw exception when opening an Aspose.Cells encrypted workbook with wrong password
// Description: Shows how to create a password‑protected workbook, save it, then attempt to load it with an incorrect password via LoadOptions, causing an exception, and finally confirms successful loading with the correct password.
// Keywords: Aspose.Cells | .NET | C# | encrypted workbook | wrong password | LoadOptions | exception handling | password protection | Workbook open error | invalid password Excel
// Common Searches: Aspose.Cells open encrypted workbook with wrong password | C# catch exception invalid password Aspose.Cells | LoadOptions password error Aspose.Cells .NET | how to validate Excel file encryption Aspose.Cells | exception thrown when password mismatch Aspose.Cells
// Developer Intent: Verify that loading a password‑protected workbook with an incorrect password raises an exception.
// Use Cases: Automated test to ensure encryption enforcement by catching the exception on a bad password. | Runtime validation that a user-provided password is correct before processing a protected workbook. | Debugging scenario to differentiate between authentication failures and file‑corruption errors.
// AI Prompts: Generate a C# unit test using Aspose.Cells that asserts an exception is thrown for a wrong password. | Provide a code snippet that logs the exact exception type and message when a mismatched password is supplied. | Explain how to distinguish a password‑mismatch exception from a file‑corruption exception when loading a workbook with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordTest
{
    // Shows how to create a password‑protected workbook, save it, then attempt to load it with an incorrect password via LoadOptions, causing an exception, and finally confirms successful loading with the correct password.
    class Program
    {
        static void Main()
        {
            // Step 1: Create a new workbook and add sample data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encrypted content");

            // Step 2: Set a password to encrypt the workbook
            wb.Settings.Password = "correctPassword";

            // Step 3: Save the encrypted workbook to disk
            string filePath = "encryptedWorkbook.xlsx";
            wb.Save(filePath);

            // Step 4: Attempt to open the encrypted workbook with an incorrect password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "wrongPassword";

            try
            {
                // This line should throw an exception because the password is incorrect
                Workbook wbWrong = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook opened unexpectedly with wrong password.");
            }
            catch (Exception ex)
            {
                // Expected path: an exception is thrown
                Console.WriteLine("Exception caught as expected when using wrong password:");
                Console.WriteLine(ex.Message);
            }

            // Optional: Verify that opening with the correct password works
            loadOptions.Password = "correctPassword";
            try
            {
                Workbook wbCorrect = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook opened successfully with correct password.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected exception with correct password: " + ex.Message);
            }
        }
    }
}
