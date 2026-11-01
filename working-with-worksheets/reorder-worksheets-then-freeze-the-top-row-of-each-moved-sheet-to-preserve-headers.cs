// Title: Reverse worksheet order and freeze the header row on each sheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells reverse worksheet order and keep first row frozen | Aspose.Cells .NET freeze top row after moving worksheets | programmatically reorder Excel sheets and apply FreezePanes with Aspose.Cells | how to copy worksheets in reverse order using Aspose.Cells C# | freeze panes on all worksheets after reordering with Aspose.Cells
// Tags: reverse worksheet order aspose.cells c# | freeze top row aspose.cells freezepanes | reorder excel sheets aspose.cells | apply FreezePanes to multiple worksheets | copy worksheets in reverse order aspose.cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a workbook, captures all worksheets, clears the original collection, adds copies of the sheets back in reverse order, freezes the first row on each worksheet using FreezePanes, and saves the result, handling missing files and I/O errors.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input workbook path
                string inputPath = "input.xlsx";
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the worksheets collection
                WorksheetCollection sheets = workbook.Worksheets;

                // Store current worksheets in a list
                List<Worksheet> sheetList = new List<Worksheet>();
                foreach (Worksheet sheet in sheets)
                {
                    sheetList.Add(sheet);
                }

                // Clear existing worksheets
                sheets.Clear();

                // Reorder worksheets (example: reverse order) using sheet names
                for (int i = sheetList.Count - 1; i >= 0; i--)
                {
                    // AddCopy expects the name of the worksheet to copy
                    sheets.AddCopy(sheetList[i].Name);
                }

                // Freeze the top row of each sheet
                foreach (Worksheet sheet in sheets)
                {
                    try
                    {
                        // Freeze first row (row index 1) and no columns
                        sheet.FreezePanes(1, 0, 1, 0);
                    }
                    catch (Exception freezeEx)
                    {
                        Console.WriteLine($"Failed to freeze panes on sheet '{sheet.Name}': {freezeEx.Message}");
                    }
                }

                // Save the modified workbook
                string outputPath = "output.xlsx";
                try
                {
                    // Ensure the output directory exists
                    string? outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
