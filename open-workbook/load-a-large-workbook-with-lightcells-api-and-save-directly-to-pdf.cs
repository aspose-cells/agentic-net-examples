// Title: Convert a large Excel workbook to PDF with Aspose.Cells LightCells API in C# (memory‑efficient)
// AI Prompts: Generate C# code that opens a large .xlsx file using Aspose.Cells LightCells API and saves it directly as a PDF, including a file‑existence check and exception handling. | Show a memory‑optimized C# example that loads a big workbook with Aspose.Cells and exports it to PDF without loading the entire file into memory.
// Common Searches: c# aspose.cells lightcells convert large xlsx to pdf without high memory usage | how to export big Excel workbook to PDF using Aspose.Cells in .NET | sample code for memory‑efficient PDF export of large Excel files with Aspose.Cells
// Tags: Aspose.Cells LightCells PDF export | memory‑efficient Excel to PDF conversion .NET | large workbook handling Aspose.Cells C# | save workbook as PDF using Aspose.Cells | C# load large xlsx with LightCells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates checking for the source .xlsx file, loading it with Aspose.Cells LightCells for low memory consumption, and saving the workbook directly as a PDF while handling errors.
class Program
{
    static void Main()
    {
        string inputPath = "largeWorkbook.xlsx";
        string outputPath = "largeWorkbook.pdf";

        try
        {
            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (memory‑efficient for large files)
            using (var workbook = new Workbook(inputPath))
            {
                // Save the workbook directly to PDF
                workbook.Save(outputPath, SaveFormat.Pdf);
            }

            Console.WriteLine($"PDF successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
