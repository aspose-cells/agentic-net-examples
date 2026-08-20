// Title: Change the password of an encrypted Excel workbook with Aspose.Cells for .NET (C#)
// Description: Loads a password‑protected .xlsx using LoadOptions, assigns a new password via workbook.Settings.Password, and saves the file. Includes error handling for CellsException and generic exceptions.
// Keywords: Aspose.Cells | C# | change Excel password | update workbook password | encrypted Excel file | LoadOptions password | Workbook.Settings.Password | re‑encrypt Excel workbook | password rotation | Excel file security
// Common Searches: how to change password of encrypted Excel using Aspose.Cells | Aspose.Cells replace workbook password C# | update Excel file password programmatically .NET | re‑encrypt Excel workbook with new password Aspose | batch change Excel passwords Aspose.Cells
// Developer Intent: Replace an existing workbook password with a stronger one programmatically.
// Use Cases: Secure legacy Excel files by updating their passwords to meet current policy requirements. | Automate bulk password rotation for multiple workbooks in a scheduled job. | Integrate password update into a CI/CD pipeline to re‑encrypt artifacts after deployment.
// AI Prompts: Generate C# code that opens an encrypted Excel workbook with Aspose.Cells, changes its password, and saves the result. | Explain how to catch and process CellsException when modifying a protected workbook's password. | Create a reusable method that takes input path, old password, new password, and output path to re‑encrypt an Excel file using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookPasswordUpdater
{
    // Loads a password‑protected .xlsx using LoadOptions, assigns a new password via workbook.Settings.Password, and saves the file. Includes error handling for CellsException and generic exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the source and destination workbooks
            string inputFile = "encrypted_workbook.xlsx";
            string outputFile = "encrypted_workbook_newpwd.xlsx";

            // Old (current) password and the new stronger password
            string oldPassword = "oldPassword123";
            string newPassword = "NewStrongPassword!2026";

            try
            {
                // Ensure the input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook using the old password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = oldPassword
                };
                Workbook workbook = new Workbook(inputFile, loadOptions);

                // Update the workbook encryption password
                workbook.Settings.Password = newPassword;

                // Save the workbook with the new password
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
