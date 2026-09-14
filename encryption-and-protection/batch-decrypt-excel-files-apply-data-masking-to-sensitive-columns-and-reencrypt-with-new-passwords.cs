// Title: Batch decrypt password‑protected Excel workbooks, mask selected columns, and re‑encrypt with new passwords using Aspose.Cells for .NET
// AI Prompts: Write a C# program that iterates over all .xlsx files in a directory, opens each workbook with a supplied password via Aspose.Cells LoadOptions, replaces the values in given column indexes with asterisks, assigns a new password to the workbook, and saves the masked file to an output folder. | Generate C# code that loads password‑protected Excel files, applies column‑wise data masking while preserving string length, updates Workbook.Settings.Password, and writes the modified workbooks as new .xlsx files encrypted with a different password using Aspose.Cells.
// Common Searches: how to programmatically decrypt multiple password protected Excel files with Aspose.Cells | c# mask sensitive data in specific columns of encrypted .xlsx files | change workbook password after editing cells using Aspose.Cells | batch process Excel files: load, mask, and re‑encrypt with new password in .NET
// Tags: load protected workbook with LoadOptions Aspose.Cells | column data masking in Excel using Aspose.Cells | re‑encrypt workbook with new password SaveFormat.Xlsx | batch processing multiple .xlsx files Aspose.Cells | update workbook Settings.Password C#

using System;
using System.IO;
using Aspose.Cells;

namespace ExcelBatchDecryptMaskEncrypt
{
    // The solution scans a source folder for .xlsx files, opens each password‑protected workbook using LoadOptions, masks the values in specified column indexes by replacing them with asterisks, sets a new workbook password via Settings.Password, and saves the masked workbooks to an output folder as new encrypted .xlsx files.
    class Program
    {
        static void Main(string[] args)
        {
            // Input parameters
            string inputFolder = @"C:\InputExcelFiles";      // Folder containing encrypted Excel files
            string outputFolder = @"C:\OutputExcelFiles";    // Folder to save re‑encrypted files
            string oldPassword = "oldPassword123";           // Current password for the files
            string newPassword = "newPassword456";           // New password to apply after masking
            // Zero‑based column indexes that contain sensitive data (e.g., 0 = A, 2 = C)
            int[] sensitiveColumns = new int[] { 1, 3 };

            // Ensure input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each .xlsx file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook with the existing password
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        Password = oldPassword
                    };
                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // Apply data masking to each worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Determine the used range to limit iteration
                        Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                        if (usedRange == null)
                            continue; // Skip empty sheets

                        int startRow = usedRange.FirstRow;
                        int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                        int startCol = usedRange.FirstColumn;
                        int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                        // Iterate over rows
                        for (int row = startRow; row <= endRow; row++)
                        {
                            foreach (int col in sensitiveColumns)
                            {
                                // Ensure column is within the used range
                                if (col >= startCol && col <= endCol)
                                {
                                    Cell cell = sheet.Cells[row, col];
                                    // Simple masking: replace the original value with asterisks preserving length
                                    if (cell?.Value != null)
                                    {
                                        string original = cell.Value.ToString();
                                        string masked = new string('*', original.Length);
                                        cell.PutValue(masked);
                                    }
                                }
                            }
                        }
                    }

                    // Set the new password for the workbook
                    workbook.Settings.Password = newPassword;

                    // Build output file path
                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    string outputPath = Path.Combine(outputFolder, $"{fileName}_masked.xlsx");

                    // Save the workbook with the new password
                    workbook.Save(outputPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Processed and saved: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
