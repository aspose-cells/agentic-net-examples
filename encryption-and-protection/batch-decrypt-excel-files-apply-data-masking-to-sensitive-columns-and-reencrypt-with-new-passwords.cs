// Title: C# batch decrypt, mask, and re‑encrypt Excel files with Aspose.Cells
// Description: Scans a folder for password‑protected Excel workbooks, opens them with Aspose.Cells, masks predefined columns (e.g., B and D) using a placeholder, applies a new password, and saves the secured files to an output directory.
// Keywords: Aspose.Cells | C# | .NET | batch decrypt Excel | password protected workbook | data masking | column masking | re‑encrypt Excel | Excel automation | GDPR compliance | secure Excel processing
// Common Searches: open password protected Excel with Aspose.Cells C# | mask sensitive columns in encrypted workbook | batch re‑encrypt Excel files .NET | replace cell values with **** in Excel | automate Excel data sanitization Aspose
// Developer Intent: Decrypt each protected workbook, replace values in selected columns with a mask, then protect the file with a new password.
// Use Cases: Automate GDPR‑oriented data sanitization by masking personal identifiers in a batch of encrypted spreadsheets before archiving. | Nightly financial reporting job that removes confidential account numbers from protected workbooks while re‑applying a corporate password. | Migrate legacy encrypted Excel files to a new security policy by batch‑processing them: decrypt, cleanse, and re‑encrypt with a unified password.
// AI Prompts: Write C# code using Aspose.Cells to open a password‑protected workbook, mask columns B and D with "****", set a new password, and save the file. | Suggest robust error‑handling and logging strategies for batch processing of encrypted Excel files with Aspose.Cells in .NET. | Explain how to externalize column indexes to a JSON config file and apply dynamic masking using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchProcessing
{
    // Scans a folder for password‑protected Excel workbooks, opens them with Aspose.Cells, masks predefined columns (e.g., B and D) using a placeholder, applies a new password, and saves the secured files to an output directory.
    class BatchDecryptMaskEncrypt
    {
        static void Main()
        {
            // Folder containing encrypted Excel files
            string inputFolder = @"C:\Input";
            // Folder where re‑encrypted files will be saved
            string outputFolder = @"C:\Output";

            // Old password used to open the files
            string oldPassword = "oldPass";
            // New password to protect the files after processing
            string newPassword = "newPass";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process all Excel files in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
            {
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm")
                    continue; // Skip non‑Excel files

                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook using the old password
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                    {
                        Password = oldPassword
                    };
                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // Define the zero‑based indexes of columns that contain sensitive data (e.g., B and D)
                    int[] sensitiveColumns = new int[] { 1, 3 };

                    // Apply masking to each sensitive column in the first worksheet
                    Worksheet sheet = workbook.Worksheets[0];
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow; // last row that contains data

                    foreach (int col in sensitiveColumns)
                    {
                        for (int row = 0; row <= maxRow; row++)
                        {
                            // Replace the original value with a masked placeholder
                            cells[row, col].PutValue("****");
                        }
                    }

                    // Set a new password for the workbook
                    workbook.Settings.Password = newPassword;

                    // Save the processed workbook to the output folder
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                    workbook.Save(outputPath);
                    Console.WriteLine($"Processed and saved: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
    }
}
