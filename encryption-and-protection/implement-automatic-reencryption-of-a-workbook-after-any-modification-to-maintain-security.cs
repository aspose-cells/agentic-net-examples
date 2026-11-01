// Title: Automatically re‑encrypt an Excel workbook with a new password after editing using Aspose.Cells for .NET
// Common Searches: how to re‑encrypt an encrypted Excel file after modifying it with Aspose.Cells C# | Aspose.Cells change workbook password after edit and save as new file | C# load password‑protected XLSX, edit cells, and save with a different password using Aspose.Cells | programmatically update encryption type and password of an Excel workbook in .NET | save modified encrypted workbook with new password using Aspose.Cells API
// Tags: re‑encrypt workbook Aspose.Cells C# | update Excel password programmatically .NET | Aspose.Cells workbook encryption settings | save encrypted XLSX with new password | change encryption type AES256 Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving; // Retained for potential future use

namespace Example
{
    // The example loads a password‑protected XLSX file with the current password, makes a cell change, assigns a new password (and optionally a different encryption type) via wb.Settings.Password, and saves the workbook as a newly encrypted file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "EncryptedInput.xlsx";
            const string outputPath = "EncryptedOutput.xlsx";
            const string currentPassword = "CurrentPassword";
            const string newPassword = "NewSecurePassword";

            try
            {
                // Verify that the input file exists before attempting to load it.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the encrypted workbook using the current password.
                var loadOptions = new LoadOptions
                {
                    Password = currentPassword
                };
                var workbook = new Workbook(inputPath, loadOptions);

                // ----- Begin modifications -----
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Updated value");
                // Additional modifications can be placed here.
                // ----- End modifications -----

                // Re‑encrypt and save the workbook with the new password.
                ReEncryptAndSave(workbook, newPassword, outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method that applies encryption settings and saves the workbook.
        static void ReEncryptAndSave(Workbook wb, string password, string outputPath)
        {
            try
            {
                // Apply new password and optional encryption type.
                wb.Settings.Password = password;
                // wb.Settings.EncryptionType = EncryptionType.AES256; // Uncomment to specify AES256.

                // Save the workbook in XLSX format.
                wb.Save(outputPath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
                throw;
            }
        }
    }
}
