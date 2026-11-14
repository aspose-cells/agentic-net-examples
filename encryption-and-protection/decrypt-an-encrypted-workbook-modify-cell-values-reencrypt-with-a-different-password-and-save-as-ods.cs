// Title: Decrypt a password‑protected XLSX workbook, edit a cell, re‑encrypt with a new password, and save as ODS using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load a password‑protected XLSX workbook | How to modify a cell and export to ODS with a new password using Aspose.Cells | Change the open password of an Excel file with Aspose.Cells .NET | Decrypt an encrypted Excel file and re‑encrypt it with a different password in C# | Save a workbook as ODS while applying password protection with Aspose.Cells
// Tags: load encrypted XLSX using LoadOptions Aspose.Cells | update cell value Aspose.Cells C# | set new open password Aspose.Cells workbook | save workbook as ODS with password Aspose.Cells | re‑encrypt Excel file Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // The example checks for an encrypted XLSX file, opens it with the original password via LoadOptions, changes cell A1 in the first worksheet, assigns a new open password, and saves the workbook as an ODS file with the new password applied.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths and passwords
            string sourcePath = "encrypted_input.xlsx";
            string oldPassword = "oldPassword";
            string newPassword = "newPassword";
            string outputPath = "modified_output.ods";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the encrypted workbook with the old password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = oldPassword
                };
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Modify cell values (example: change A1 in the first worksheet)
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("New Value");

                // Re‑encrypt with a new password (open password)
                workbook.Settings.Password = newPassword;
                // Write protection password is optional; omitted if not supported by the current library version

                // Save as ODS with the new password applied
                OdsSaveOptions saveOptions = new OdsSaveOptions(SaveFormat.Ods);
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine("Workbook decrypted, modified, re‑encrypted, and saved as ODS successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
