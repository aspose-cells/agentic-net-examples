// Title: Re‑encrypt a password‑protected XLSX workbook with a stronger AES‑256 password using Aspose.Cells for .NET
// AI Prompts: Load an existing password‑protected XLSX file with LoadOptions, assign a new password to Workbook.Settings.Password (which applies AES‑256 encryption by default), and save the workbook to a new file. | Open an encrypted Excel workbook using its current password, replace the password to trigger a higher‑strength encryption algorithm, and write the upgraded file back to disk with Aspose.Cells.
// Common Searches: Aspose.Cells C# load password protected XLSX and change to AES-256 | How to re‑encrypt an existing Excel workbook with a new password using Aspose.Cells | C# upgrade encryption algorithm of an encrypted workbook with Aspose.Cells | Replace old password with new one and apply stronger encryption in Aspose.Cells
// Tags: aspnet cells load encrypted workbook | aspnet cells change workbook password | aspnet cells aes-256 encryption upgrade | aspnet cells save workbook with new password | excel file re-encrypt aspnet cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example verifies the presence of an encrypted_input.xlsx file, loads it with the original password via LoadOptions, sets a new password on workbook.Settings.Password (which defaults to AES‑256 encryption), ensures the output directory exists, and saves the upgraded workbook as encrypted_upgraded.xlsx, handling any exceptions that may arise.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths and passwords
                string inputPath = "encrypted_input.xlsx";
                string currentPassword = "oldPassword";
                string newPassword = "newPassword";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the encrypted workbook using the current password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = currentPassword
                };
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Set new password for the workbook (default AES‑256 encryption)
                workbook.Settings.Password = newPassword;

                // Prepare output path
                string outputPath = "encrypted_upgraded.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the new password
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved with upgraded encryption to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
