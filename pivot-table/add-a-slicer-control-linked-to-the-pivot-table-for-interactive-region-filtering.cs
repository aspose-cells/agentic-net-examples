// Title: Create a Region slicer linked to a pivot table in an existing Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to add a slicer for the Region field to a pivot table with Aspose.Cells C# | Aspose.Cells C# example linking slicer to pivot table field | add interactive slicer control to existing Excel workbook using Aspose.Cells .NET | C# code to place a slicer at specific cell location in worksheet with Aspose.Cells
// Tags: Aspose.Cells insert slicer into pivot table | C# create region slicer for Excel | Aspose.Cells slicer linked to pivot field | Excel workbook slicer placement .NET | interactive pivot filtering with slicer

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;      // Required for PivotTable class
using Aspose.Cells.Slicers;   // Required for Slicer functionality

namespace AsposeCellsSlicerExample
{
    // The example loads an existing workbook, checks for a worksheet named "Sheet1" and a pivot table, then inserts a slicer linked to the "Region" field at cell E5, sets its name and caption, and saves the modified file as a new Excel workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the worksheet that contains the pivot table (named "Sheet1")
                Worksheet worksheet = workbook.Worksheets["Sheet1"];
                if (worksheet == null)
                {
                    Console.WriteLine("Worksheet \"Sheet1\" not found in the workbook.");
                    return;
                }

                // Ensure the worksheet has at least one pivot table
                if (worksheet.PivotTables == null || worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the worksheet.");
                    return;
                }

                // Assume the first pivot table on the sheet is the target
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Add a slicer linked to the "Region" field of the pivot table
                // Note: In this API version the parameters are (pivotTable, firstRow, firstColumn, fieldName)
                int slicerIndex = worksheet.Slicers.Add(pivotTable, 5, 5, "Region");

                // Retrieve the slicer object to customize it if needed
                Slicer slicer = worksheet.Slicers[slicerIndex];
                slicer.Name = "RegionSlicer";
                slicer.Caption = "Region Filter";

                // Save the workbook with the new slicer
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
