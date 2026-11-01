// Title: Save a large Excel workbook as XLS in C# with Aspose.Cells using a custom temporary folder to avoid RAM exhaustion
// Common Searches: C# Aspose.Cells save large workbook as XLS using temporary folder to reduce memory usage | How to configure Aspose.Cells MemoryPreference and custom temp directory for big Excel files in .NET | Saving 20,000 rows to XLS with Aspose.Cells without out-of-memory errors | Set Aspose.Cells temporary folder path for large workbook export in C# | Aspose.Cells large file handling memory preference example C#
// Tags: Aspose.Cells MemoryPreference configuration | custom temporary folder for Aspose.Cells | save workbook as XLS in C# | large Excel file memory management with Aspose.Cells | C# large workbook export using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, fills 20,000 rows, creates a custom temporary directory (C:\AsposeTemp), sets workbook.Settings.MemorySetting to MemoryPreference so Aspose.Cells writes temporary files instead of keeping everything in RAM, and saves the file as LargeWorkbook.xls.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate the workbook with sample data (optional, for demonstration)
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 20000; row++)
            {
                sheet.Cells[row, 0].PutValue($"Row {row}");
                sheet.Cells[row, 1].PutValue(row);
            }

            // Specify a temporary folder for Aspose.Cells to use when handling large files
            string tempFolderPath = @"C:\AsposeTemp";
            if (!Directory.Exists(tempFolderPath))
            {
                Directory.CreateDirectory(tempFolderPath);
            }

            // Configure the workbook to use temporary files instead of keeping everything in RAM
            workbook.Settings.MemorySetting = MemorySetting.MemoryPreference; // use temp files

            // Note: In recent Aspose.Cells versions the TempFolder property is not exposed.
            // Aspose.Cells will use the system temporary folder or the folder set via environment variables.

            // Save the workbook as an XLS file (format inferred from extension)
            string outputPath = "LargeWorkbook.xls";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
