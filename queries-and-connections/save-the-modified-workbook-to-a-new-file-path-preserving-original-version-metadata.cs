// Title: Save a modified Excel workbook to a new file path without losing original version metadata using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an existing .xlsx workbook with Aspose.Cells, applies any changes, and saves it to a different location while keeping the original file's version metadata intact. | Create a C# snippet that verifies the source Excel file exists, updates a cell, and writes the updated workbook to a new path using Workbook.Save with SaveFormat.Xlsx, ensuring metadata such as version information is preserved.
// Common Searches: Aspose.Cells keep original file properties when saving a modified workbook | C# save a copy of an Excel file with changes without altering its version metadata | preserve workbook metadata after editing with Aspose.Cells .NET | how to write modified Excel workbook to new location while retaining creation date using Aspose.Cells
// Tags: Aspose.Cells workbook.Save with SaveFormat.Xlsx | preserve file version metadata Aspose.Cells | C# copy and modify Excel workbook Aspose.Cells | save modified workbook to different path .NET | check file existence before loading Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing .xlsx file with Aspose.Cells, optionally modifies its content, and saves the workbook to a new file path using the Xlsx format. It includes a file‑existence check and exception handling, ensuring that the original workbook's version metadata remains unchanged.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = @"C:\Path\To\OriginalWorkbook.xlsx";
        string outputPath = @"C:\Path\To\ModifiedWorkbook.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the original workbook
            Workbook workbook = new Workbook(inputPath);

            // ---- Perform any modifications to the workbook here ----
            // Example modification:
            // workbook.Worksheets[0].Cells["A1"].PutValue("Modified");

            // Save the modified workbook to the specified output path
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
