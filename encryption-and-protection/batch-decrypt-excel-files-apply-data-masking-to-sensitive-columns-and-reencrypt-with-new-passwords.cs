// Title: C# – Batch Decrypt, Mask Columns, and Re‑Encrypt Excel Files with Aspose.Cells
// Description: A C# utility that scans a folder, detects password‑protected .xlsx files, opens them with the old password, replaces values in specified columns with a mask, applies a new password, and saves the secured workbooks to an output directory.
// Keywords: Aspose.Cells | C# Excel encryption | batch decrypt Excel | data masking Excel | re‑encrypt workbook | password protected workbook | bulk Excel processing | detect encrypted Excel file | column redaction | Excel security automation
// Common Searches: how to decrypt multiple Excel files with Aspose.Cells | C# code to mask sensitive columns in encrypted workbooks | batch re‑encrypt Excel files with a new password .NET | detect and open password protected .xlsx using Aspose.Cells | automate GDPR redaction in Excel spreadsheets C#
// Developer Intent: Open each workbook, mask defined columns, and save it encrypted with a new password.
// Use Cases: Sanitize personal data in a collection of password‑protected reports before sharing. | Migrate legacy encrypted spreadsheets to a new security policy while obscuring confidential fields. | Automate GDPR‑compliant redaction of specific columns across many Excel files. | Create a secure archive of Excel workbooks with updated passwords after data masking.
// AI Prompts: Generate C# code using Aspose.Cells to bulk open encrypted .xlsx files, replace values in columns B and D with "*****", and save them with a new password. | Show an Aspose.Cells example that detects if an Excel file is password protected, loads it with a given password, masks selected columns, and re‑encrypts it with another password. | Explain how to extend the program to log processed file names, masking actions, and errors to a CSV report. | Provide guidance on customizing the mask placeholder and adding column selection via a configuration file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// A C# utility that scans a folder, detects password‑protected .xlsx files, opens them with the old password, replaces values in specified columns with a mask, applies a new password, and saves the secured workbooks to an output directory.
class Program
{
    // Adjust these values as needed
    private static readonly string InputFolder = @"C:\InputExcelFiles";
    private static readonly string OutputFolder = @"C:\OutputExcelFiles";
    private static readonly string OldPassword = "oldPass123";
    private static readonly string NewPassword = "newPass456";

    // Zero‑based column indexes that contain sensitive data (e.g., B and D columns)
    private static readonly int[] SensitiveColumns = { 1, 3 };

    static void Main()
    {
        // Ensure output directory exists
        Directory.CreateDirectory(OutputFolder);

        // Process each Excel file in the input folder
        foreach (string filePath in Directory.GetFiles(InputFolder, "*.xlsx"))
        {
            try
            {
                // Detect whether the file is encrypted
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                Workbook workbook;

                if (formatInfo.IsEncrypted)
                {
                    // Load encrypted workbook using the old password
                    LoadOptions loadOptions = new LoadOptions();
                    loadOptions.Password = OldPassword;
                    workbook = new Workbook(filePath, loadOptions);
                }
                else
                {
                    // Load unencrypted workbook
                    workbook = new Workbook(filePath);
                }

                // Apply simple masking to each sensitive column in the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                int totalRows = sheet.Cells.Rows.Count;

                for (int row = 0; row < totalRows; row++)
                {
                    foreach (int col in SensitiveColumns)
                    {
                        var cell = sheet.Cells[row, col];
                        if (cell != null && cell.Value != null)
                        {
                            // Replace the original value with a masked placeholder
                            cell.PutValue("*****");
                        }
                    }
                }

                // Set a new password for the workbook
                workbook.Settings.Password = NewPassword;

                // Save the re‑encrypted workbook to the output folder
                string outputPath = Path.Combine(OutputFolder, Path.GetFileName(filePath));
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
