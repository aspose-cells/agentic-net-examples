using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class BatchDecryptMaskEncrypt
{
    static void Main()
    {
        // Folder containing the source Excel files
        string sourceFolder = @"C:\InputExcelFiles";
        // Folder where the processed files will be saved
        string outputFolder = @"C:\OutputExcelFiles";

        // Old password used to open encrypted files (if any)
        string oldPassword = "oldPassword123";
        // New password to protect the processed files
        string newPassword = "newPassword456";

        // Zero‑based column indexes that contain sensitive data (e.g., 1 = B, 3 = D)
        int[] sensitiveColumns = { 1, 3 };

        // Ensure the output folder exists
        Directory.CreateDirectory(outputFolder);

        // Verify the source folder exists before enumerating files
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder not found: {sourceFolder}");
            return;
        }

        // Process each Excel file in the source folder
        foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
        {
            try
            {
                // Guard against missing files (should not happen after GetFiles, but kept for safety)
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Detect file format and encryption status
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                Workbook workbook;

                // Load the workbook with appropriate password if it is encrypted
                if (formatInfo.IsEncrypted)
                {
                    LoadOptions loadOptions = new LoadOptions
                    {
                        Password = oldPassword
                    };
                    workbook = new Workbook(filePath, loadOptions);
                }
                else
                {
                    workbook = new Workbook(filePath);
                }

                // Apply data masking on each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int lastRow = sheet.Cells.MaxDataRow; // last row with data

                    for (int row = 0; row <= lastRow; row++)
                    {
                        foreach (int colIndex in sensitiveColumns)
                        {
                            Cell cell = sheet.Cells[row, colIndex];
                            if (cell.Value != null && !string.IsNullOrWhiteSpace(cell.StringValue))
                            {
                                cell.PutValue("****");
                            }
                        }
                    }
                }

                // Set the new password for the workbook
                workbook.Settings.Password = newPassword;

                // Build the output file path
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                // Save the processed workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}