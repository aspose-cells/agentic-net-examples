// Title: Load an XLSX workbook from a file path and strip all embedded charts using Aspose.Cells for .NET
// Common Searches: how to open an Excel .xlsx file and delete all charts with Aspose.Cells C# | Aspose.Cells remove chart objects after loading workbook in .NET | C# code to clear worksheet charts when reading an XLSX file | load workbook without embedded drawings using Aspose.Cells library
// Tags: load xlsx workbook Aspose.Cells | clear worksheet charts Aspose.Cells | remove embedded Excel charts C# | save workbook without charts Aspose.Cells | chart collection clear Aspose.Cells API

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the source file, loads the XLSX workbook with Aspose.Cells, iterates through each worksheet to clear the Charts and Shapes collections, ensures the output directory exists, and saves a new workbook that no longer contains any embedded chart objects.
class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = @"C:\Path\To\Your\File.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the workbook (charts will be removed after loading)
            Workbook workbook = new Workbook(sourcePath);

            // Remove all chart objects from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Charts.Clear();   // Delete charts
                sheet.Shapes.Clear();   // Optionally delete other drawing objects
            }

            // Path for the output workbook without charts
            string outputPath = @"C:\Path\To\Your\OutputWithoutCharts.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved without charts to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
