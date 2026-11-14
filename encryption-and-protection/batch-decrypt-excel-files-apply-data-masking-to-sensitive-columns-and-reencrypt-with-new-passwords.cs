// Title: How to batch decrypt password‑protected Excel files, mask sensitive columns, and re‑encrypt them with new passwords using Aspose.Cells for .NET
// Common Searches: C# code to decrypt multiple encrypted .xlsx files, mask PII columns, and save with a new password using Aspose.Cells | Aspose.Cells load workbook with password and apply column masking in a folder of Excel files | Change password of a batch of Excel workbooks after masking data in specific columns in C# | Programmatically mask string, numeric, and date cells in protected Excel files with Aspose.Cells | How to process all .xls and .xlsx files in a directory to re‑encrypt after data masking in .NET
// Tags: batch decrypt Excel workbooks using Aspose.Cells | apply data masking to specific columns in Excel with Aspose.Cells | set new workbook password programmatically in C# | process multiple .xls and .xlsx files in a directory Aspose.Cells | load encrypted workbook with LoadOptions Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelBatchDecryptMaskEncrypt
{
    // The sample scans an input folder for .xls and .xlsx files, loads each workbook with the old password via LoadOptions, masks string, numeric, and date values in the specified column indices across all worksheets, assigns a new password, and saves the masked workbook as a new .xlsx file in an output folder.
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // Input folder containing encrypted Excel files
            string inputFolder = @"C:\InputExcelFiles";
            // Output folder for re‑encrypted files
            string outputFolder = @"C:\OutputExcelFiles";
            // Old password used to decrypt the files
            string oldPassword = "oldPassword123";
            // New password to encrypt the files after masking
            string newPassword = "newPassword456";
            // Zero‑based column indices that contain sensitive data (e.g., 0 = A, 2 = C)
            int[] sensitiveColumns = new int[] { 1, 3 }; // columns B and D

            ProcessFolder(inputFolder, outputFolder, oldPassword, newPassword, sensitiveColumns);
        }

        static void ProcessFolder(string inputFolder, string outputFolder, string oldPwd, string newPwd, int[] sensitiveCols)
        {
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all Excel files (both .xlsx and .xls) in the input folder
            string[] files = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in files)
            {
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xlsx" && ext != ".xls")
                    continue; // skip non‑Excel files

                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook with the old password
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                    {
                        Password = oldPwd
                    };
                    Workbook wb = new Workbook(filePath, loadOptions);

                    // Apply data masking to the specified columns
                    MaskSensitiveColumns(wb, sensitiveCols);

                    // Set new password for the workbook
                    wb.Settings.Password = newPwd;

                    // Build output file path
                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    string outputPath = Path.Combine(outputFolder, fileName + "_masked.xlsx");

                    // Save the workbook (as .xlsx) with the new password
                    wb.Save(outputPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Processed: {filePath} -> {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }

        static void MaskSensitiveColumns(Workbook wb, int[] columnIndices)
        {
            foreach (Worksheet sheet in wb.Worksheets)
            {
                // Determine the used range to limit iteration
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // If the sheet is empty, skip it
                if (usedRange == null)
                    continue;

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    foreach (int col in columnIndices)
                    {
                        // Ensure column index is within the used range
                        if (col < startCol || col > endCol)
                            continue;

                        Cell cell = sheet.Cells[row, col];
                        if (cell.Type == CellValueType.IsString)
                        {
                            string original = cell.StringValue;
                            // Simple masking: replace each character with '*'
                            string masked = new string('*', original.Length);
                            cell.PutValue(masked);
                        }
                        else if (cell.Type == CellValueType.IsNumeric)
                        {
                            // For numeric data, replace with zero
                            cell.PutValue(0);
                        }
                        else if (cell.Type == CellValueType.IsDateTime)
                        {
                            // For dates, replace with a default date (e.g., 01/01/1900)
                            cell.PutValue(new DateTime(1900, 1, 1));
                        }
                        // Other types (Boolean, Error, etc.) can be handled as needed
                    }
                }
            }
        }
    }
}
