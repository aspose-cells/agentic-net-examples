// Title: Copy all cells, styles, and formatting from one worksheet to another within the same workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy worksheet contents to another sheet without creating a new workbook | How to duplicate a sheet with formatting using Aspose.Cells for .NET | Copy data and styles between worksheets in the same Excel file with Aspose.Cells | C# Aspose.Cells copy Sheet1 to Sheet2 preserving cell formatting | Transfer entire worksheet to another tab using Aspose.Cells Worksheet.Copy method
// Tags: worksheet copy Aspose.Cells .NET | duplicate sheet preserving formatting C# | copy cells and styles between worksheets | Aspose.Cells copy sheet without new workbook | transfer worksheet data within same workbook

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, ensures source (Sheet1) and destination (Sheet2) worksheets exist, uses the Worksheet.Copy method to duplicate all cells, styles, and other sheet contents from the source to the destination, and saves the updated workbook, handling errors gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string outputPath = "output.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load workbook containing both source and destination worksheets
                Workbook workbook = new Workbook(sourcePath);

                // Get the source worksheet (e.g., "Sheet1")
                Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
                if (sourceSheet == null)
                {
                    Console.WriteLine("Source worksheet 'Sheet1' not found.");
                    return;
                }

                // Ensure the destination worksheet exists (e.g., "Sheet2")
                Worksheet destinationSheet = workbook.Worksheets["Sheet2"];
                if (destinationSheet == null)
                {
                    destinationSheet = workbook.Worksheets.Add("Sheet2");
                }

                // Copy all cells, styles, and other sheet contents from sourceSheet to destinationSheet
                sourceSheet.Copy(destinationSheet);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
