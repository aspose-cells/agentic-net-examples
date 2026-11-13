// Title: Save a modified Excel workbook to a new file while keeping the original format using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook to different location preserving original Excel format | How to map source file extension to SaveFormat in Aspose.Cells | Create missing directory before saving an Excel file with Aspose.Cells in C# | Load, edit, and re‑save an existing .xlsx file without changing its version using Aspose.Cells | Save modified workbook to new path while retaining original version metadata Aspose.Cells
// Tags: save workbook with original Excel version Aspose.Cells | map file extension to Aspose.Cells SaveFormat | ensure output directory exists before saving | edit cell and re‑save workbook Aspose.Cells | preserve Excel metadata during save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, updates cell A1, determines the correct SaveFormat based on the original file's extension, creates the target folder if it does not exist, and saves the modified workbook to a new location while preserving the original version and format metadata.
    class Program
    {
        static void Main()
        {
            // Paths for the original and modified workbooks
            string originalFilePath = @"C:\Path\To\OriginalWorkbook.xlsx";
            string newFilePath = @"C:\Path\To\ModifiedWorkbook.xlsx";

            try
            {
                // Verify that the original file exists
                if (!File.Exists(originalFilePath))
                {
                    Console.WriteLine($"Original file not found: {originalFilePath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(originalFilePath);

                // ----- Modify the workbook as needed -----
                // Example modification: write a value to cell A1 of the first worksheet
                workbook.Worksheets[0].Cells["A1"].PutValue("Modified");

                // Determine the appropriate SaveFormat based on the original file extension
                SaveFormat saveFormat = GetSaveFormatFromExtension(originalFilePath);

                // Ensure the target directory exists
                string targetDir = Path.GetDirectoryName(newFilePath);
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }

                // Save the modified workbook
                workbook.Save(newFilePath, saveFormat);
                Console.WriteLine($"Workbook saved successfully to: {newFilePath}");
            }
            catch (Exception ex)
            {
                // Runtime safety: capture and display any errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to map file extension to Aspose.Cells SaveFormat
        private static SaveFormat GetSaveFormatFromExtension(string filePath)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension == ".xls" ? SaveFormat.Excel97To2003 : SaveFormat.Xlsx;
        }
    }
}
