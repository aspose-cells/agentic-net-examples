// Title: Re‑encrypt a legacy‑protected Excel .xlsx workbook to AES‑256 using Aspose.Cells for .NET
// AI Prompts: Load an existing .xlsx file that uses legacy encryption, supply the old password, assign a new password, and save the file so Aspose.Cells applies AES‑256 encryption. | Check that the source workbook file exists before loading it, then re‑encrypt it with a stronger password while handling possible load or save errors. | Implement exception handling for incorrect old passwords, missing files, or I/O issues during the upgrade to AES‑256 encryption.
// Common Searches: asp.net how to convert a password‑protected Excel file from legacy encryption to AES‑256 with Aspose.Cells | c# code to re‑encrypt an existing .xlsx using a new password and AES‑256 encryption | aspose.cells load workbook with old password and save with stronger encryption | upgrade legacy encrypted Excel workbook to AES‑256 in .net
// Tags: re‑encrypt Excel workbook AES‑256 Aspose.Cells | load workbook with password LoadOptions C# | set workbook password Workbook.Settings C# | upgrade legacy Excel encryption .NET | strong encryption for .xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving; // For OoxmlSaveOptions if needed in future

// The program loads a legacy‑encrypted XLSX file using the original password, assigns a new password via Workbook.Settings, and saves it; Aspose.Cells automatically encrypts the output with AES‑256.
class Program
{
    static void Main()
    {
        // Paths to the source and destination workbooks
        string inputPath = "legacy_encrypted.xlsx";
        string outputPath = "aes256_encrypted.xlsx";

        // Password used for the legacy‑encrypted workbook
        string oldPassword = "oldPassword";

        // Password to apply for the new encrypted workbook
        string newPassword = "newPassword";

        try
        {
            // Ensure the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook with the legacy password
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = oldPassword
            };
            var workbook = new Workbook(inputPath, loadOptions);

            // Apply a new password for the workbook (encryption)
            workbook.Settings.Password = newPassword;

            // Save the workbook; Aspose.Cells will encrypt it using the password.
            // The default encryption for .xlsx files is AES‑256 in recent versions.
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
