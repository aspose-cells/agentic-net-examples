// Title: How to replace an existing password on an encrypted Excel .xlsx workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an encrypted .xlsx file with its current password via LoadOptions, assign a stronger password to Workbook.Settings.Password, and save the workbook using OoxmlSaveOptions. | Show how to re‑encrypt a password‑protected Excel workbook in C# by updating the password property and persisting the file with Aspose.Cells.
// Common Searches: Aspose.Cells change password of encrypted Excel file C# | C# update workbook password after loading with old password Aspose.Cells | re‑encrypt XLSX with new password using Aspose.Cells .NET | how to set new password when saving a protected workbook Aspose.Cells
// Tags: update workbook password Aspose.Cells | load encrypted xlsx with password Aspose.Cells | save workbook with new password OoxmlSaveOptions | change Excel file protection C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an encrypted XLSX workbook using the original password, sets a stronger password via Workbook.Settings.Password, and saves the workbook to a new file with the updated protection using Aspose.Cells for .NET.
class ChangeWorkbookPassword
{
    static void Main()
    {
        // Paths to the source and destination files
        string inputPath = "encrypted.xlsx";
        string outputPath = "encrypted_new.xlsx";

        // Old (current) password and the new stronger password
        string oldPassword = "oldPassword123";
        string newPassword = "NewStrong!Pass456";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the encrypted workbook using the old password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = oldPassword
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Update the workbook's password (this will be used when saving)
            workbook.Settings.Password = newPassword;

            // Save the workbook with the new password
            // No need to set password on OoxmlSaveOptions; the workbook's Settings handle it
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved successfully with new password to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
