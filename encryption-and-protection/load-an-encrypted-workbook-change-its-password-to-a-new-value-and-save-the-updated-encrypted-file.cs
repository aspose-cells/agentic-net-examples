// Title: C# – Change the password of an encrypted Excel workbook with Aspose.Cells
// Description: Load an encrypted .xlsx file using LoadOptions, assign a new opening password via workbook.Settings.Password, and save the workbook. The example includes directory checks and robust exception handling for Aspose.Cells and general errors.
// Keywords: Aspose.Cells C# password change | update Excel workbook encryption | load encrypted Excel file .NET | set new workbook password | Excel file re‑encryption programmatically | Workbook.Settings.Password | LoadOptions password
// Common Searches: change password of encrypted Excel file Aspose.Cells C# | load workbook with password and save with new password .NET | programmatically modify Excel workbook opening password | Aspose.Cells example for password rotation | C# code to re‑encrypt Excel workbook
// Developer Intent: Replace the existing opening password of an encrypted Excel workbook with a new one and save the file using Aspose.Cells for .NET.
// Use Cases: User updates credentials and needs the workbook re‑protected. | Organization enforces a new corporate password policy across multiple reports. | Automated routine to rotate passwords for stored Excel analytics files.
// AI Prompts: Generate C# code that opens an encrypted Excel workbook with Aspose.Cells, changes its opening password, and saves the file with the new password. | Explain how to handle CellsException when the supplied old password is incorrect while changing workbook encryption. | Show how to verify the output directory exists before saving a re‑encrypted workbook in C#.

using System;
using System.IO;
using Aspose.Cells;

namespace ChangeWorkbookPassword
{
    // Load an encrypted .xlsx file using LoadOptions, assign a new opening password via workbook.Settings.Password, and save the workbook. The example includes directory checks and robust exception handling for Aspose.Cells and general errors.
    class Program
    {
        static void Main()
        {
            // Path to the existing encrypted workbook
            string inputPath = "encrypted_input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Current password used to protect the workbook
            string oldPassword = "oldPassword123";

            // New password to set for the workbook
            string newPassword = "newPassword456";

            try
            {
                // Load the encrypted workbook using the old password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = oldPassword
                };
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Change the workbook's opening password to the new value
                workbook.Settings.Password = newPassword;

                // Define output path and ensure the directory exists
                string outputPath = "encrypted_output.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook; it will be saved with the new password
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook password changed from '{oldPassword}' to '{newPassword}' and saved to '{outputPath}'.");
            }
            catch (CellsException ex)
            {
                // Handle errors related to Aspose.Cells operations (e.g., invalid password)
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
