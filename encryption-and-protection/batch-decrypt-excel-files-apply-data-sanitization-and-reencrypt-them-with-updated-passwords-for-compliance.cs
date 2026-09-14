// Title: Batch decrypt, sanitize, and re‑encrypt Excel .xlsx files with Aspose.Cells for .NET using a common password
// AI Prompts: Write C# code that iterates through all .xlsx files in a folder, opens each with a known password via Aspose.Cells LoadOptions, replaces any cell containing "SSN" with "REDACTED", sets a new workbook password, and saves the sanitized file to an output directory. | Generate a C# loop that loads encrypted Excel workbooks, applies custom data‑redaction rules to cell values, updates the workbook's Settings.Password property, and writes the cleaned workbooks back using Aspose.Cells SaveFormat.Xlsx.
// Common Searches: how to programmatically decrypt multiple encrypted Excel files with Aspose.Cells C# | batch replace sensitive text in password protected .xlsx using Aspose.Cells | C# change workbook password after sanitizing data in encrypted Excel files | Aspose.Cells load encrypted workbook with old password and save with new password | automate redaction of SSN in protected Excel workbooks using .NET
// Tags: batch decrypt encrypted Excel workbooks Aspose.Cells | cell value redaction in protected .xlsx files | update workbook password after sanitization C# | load workbook with password LoadOptions Aspose.Cells | save workbook with new password SaveFormat.Xlsx

using System;
using System.IO;
using Aspose.Cells;

// The program scans a directory of password‑protected .xlsx files, opens each workbook with the old password using LoadOptions, iterates through all cells to replace any occurrence of "SSN" with "REDACTED", assigns a new password via workbook.Settings.Password, and saves the sanitized workbooks to a separate output folder.
class BatchDecryptSanitizeEncrypt
{
    static void Main()
    {
        // Folder containing encrypted Excel files
        string inputFolder = @"C:\EncryptedExcels";
        // Folder to save re‑encrypted sanitized files
        string outputFolder = @"C:\SanitizedExcels";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // Old password used to decrypt the files (assumed same for all files)
        string oldPassword = "OldPassword123";
        // New password to encrypt the sanitized files
        string newPassword = "NewPassword456";

        try
        {
            // Process each .xlsx file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                // Verify the file exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook with the old password
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        Password = oldPassword
                    };
                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // ----- Data Sanitization -----
                    // Example: Replace any cell containing the word "SSN" with "REDACTED"
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        var cells = sheet.Cells;
                        int maxRow = cells.MaxDataRow;
                        int maxCol = cells.MaxDataColumn;

                        for (int row = 0; row <= maxRow; row++)
                        {
                            for (int col = 0; col <= maxCol; col++)
                            {
                                var cell = cells[row, col];
                                if (cell.Type == CellValueType.IsString && cell.StringValue.Contains("SSN"))
                                {
                                    cell.PutValue("REDACTED");
                                }
                                // Add additional sanitization rules here as needed
                            }
                        }
                    }

                    // Set new password for the workbook before saving
                    workbook.Settings.Password = newPassword;

                    // Save the sanitized workbook
                    string outputFilePath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                    workbook.Save(outputFilePath, SaveFormat.Xlsx);

                    Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
