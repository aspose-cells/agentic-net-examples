// Title: C# batch decrypt password‑protected .xlsx files, mask SSN, CreditCard and Email columns, and re‑encrypt with a new password using Aspose.Cells
// Common Searches: C# Aspose.Cells load multiple encrypted .xlsx workbooks with an old password and save them with a new password | How to mask SSN and CreditCard columns in password‑protected Excel files using Aspose.Cells | Batch process Excel files: decrypt, apply data masking, and re‑encrypt in a .NET application | Aspose.Cells example for reading a protected workbook, modifying cell values, and setting a new workbook password | C# code to replace values in specific columns of encrypted Excel worksheets
// Tags: batch decrypt encrypted xlsx Aspose.Cells | mask sensitive columns Aspose.Cells | re‑encrypt workbook with new password C# | load workbook with password LoadOptions | save workbook with password SaveFormat.Xlsx | identify column indexes by header Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

// The sample scans an input folder for .xlsx files, opens each workbook with the old password via Aspose.Cells LoadOptions, detects columns named SSN, CreditCard, or Email from the header row, replaces their values with asterisks, assigns a new password to the workbook, and saves the masked, re‑encrypted file to an output folder while preserving the original file name.
class ExcelBatchProcessor
{
    // Define the folder containing the source files and the folder for the processed files
    private const string InputFolder = @"C:\Excel\Input";
    private const string OutputFolder = @"C:\Excel\Output";

    // Old password used to decrypt the source files
    private const string OldPassword = "oldPassword123";

    // New password to encrypt the processed files
    private const string NewPassword = "newPassword456";

    // List of column headers that contain sensitive data and need masking
    private static readonly string[] SensitiveColumns = { "SSN", "CreditCard", "Email" };

    static void Main()
    {
        try
        {
            // Ensure the output directory exists
            Directory.CreateDirectory(OutputFolder);

            // Verify the input directory exists
            if (!Directory.Exists(InputFolder))
            {
                Console.WriteLine($"Input folder not found: {InputFolder}");
                return;
            }

            // Get all Excel files in the input folder
            string[] excelFiles = Directory.GetFiles(InputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in excelFiles)
            {
                try
                {
                    // Verify the source file still exists
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Source file not found: {filePath}");
                        continue;
                    }

                    // Load the workbook with the old password
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        Password = OldPassword
                    };
                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // Apply data masking to each worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Determine the column indexes that need masking based on the header row (assumed to be row 0)
                        Dictionary<int, bool> columnsToMask = GetSensitiveColumnIndexes(sheet);

                        // If no sensitive columns are found, skip masking for this sheet
                        if (columnsToMask.Count == 0) continue;

                        // Iterate over rows starting from row 1 (skip header)
                        int maxRow = sheet.Cells.MaxDataRow;
                        for (int row = 1; row <= maxRow; row++)
                        {
                            foreach (int colIndex in columnsToMask.Keys)
                            {
                                Cell cell = sheet.Cells[row, colIndex];
                                if (cell != null && cell.Type != CellValueType.IsNull)
                                {
                                    // Simple masking: replace the original value with asterisks
                                    cell.PutValue("*****");
                                }
                            }
                        }
                    }

                    // Set the new password for the workbook
                    workbook.Settings.Password = NewPassword;

                    // Build the output file path (preserve original file name)
                    string outputFilePath = Path.Combine(OutputFolder, Path.GetFileName(filePath));

                    // Save the workbook with the new password
                    workbook.Save(outputFilePath, SaveFormat.Xlsx);

                    Console.WriteLine($"Processed and re‑encrypted: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    // Log or handle errors for individual files without stopping the batch
                    Console.WriteLine($"Error processing file '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Log unexpected errors that prevent the batch from starting
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }

    /// <param name="sheet">Worksheet to inspect.</param>
    /// <returns>A dictionary where the key is the column index to mask.</returns>
    private static Dictionary<int, bool> GetSensitiveColumnIndexes(Worksheet sheet)
    {
        Dictionary<int, bool> result = new Dictionary<int, bool>();
        int maxColumn = sheet.Cells.MaxDataColumn;

        for (int col = 0; col <= maxColumn; col++)
        {
            Cell headerCell = sheet.Cells[0, col];
            if (headerCell != null && headerCell.Type == CellValueType.IsString)
            {
                string headerText = headerCell.StringValue.Trim();
                foreach (string sensitiveHeader in SensitiveColumns)
                {
                    if (string.Equals(headerText, sensitiveHeader, StringComparison.OrdinalIgnoreCase))
                    {
                        result[col] = true;
                        break;
                    }
                }
            }
        }

        return result;
    }
}
