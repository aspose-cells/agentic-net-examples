// Title: Read merged cell ranges with Aspose.Cells LightCells API in C# and output their start and end coordinates per worksheet
// AI Prompts: Generate a C# console program that loads an Excel file using Aspose.Cells LightCells API, iterates every worksheet, and prints the start row/column and end row/column of each merged cell range. | Show how to use Worksheet.Cells.GetMergedAreas() in C# to retrieve merged areas and output their row/column bounds for layout analysis.
// Common Searches: how to list merged cell positions using Aspose.Cells in C# | C# Aspose.Cells LightCells get merged ranges from each worksheet | retrieve start and end rows of merged cells with Aspose.Cells API | enumerate merged areas in an Excel workbook using Aspose.Cells C# | Aspose.Cells read merged cell ranges for layout processing
// Tags: Aspose.Cells LightCells merged area enumeration | C# read merged cell coordinates Excel | Worksheet.Cells.GetMergedAreas usage | merged cell range extraction Aspose.Cells | layout analysis merged cells C#

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, iterates each worksheet, uses Worksheet.Cells.GetMergedAreas() to obtain all merged cell ranges, and writes the sheet name together with the start and end row/column indices of each range to the console.
class MergedCellReader
{
    static void Main()
    {
        // Path to the Excel file to be processed
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                string sheetName = worksheet.Name;

                // Retrieve all merged cell ranges for the current worksheet
                CellArea[] mergedRanges = worksheet.Cells.GetMergedAreas();

                // Output each merged range's start and end coordinates
                foreach (CellArea area in mergedRanges)
                {
                    Console.WriteLine(
                        $"Sheet: {sheetName}, " +
                        $"Start: (Row {area.StartRow}, Column {area.StartColumn}), " +
                        $"End: (Row {area.EndRow}, Column {area.EndColumn})");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
