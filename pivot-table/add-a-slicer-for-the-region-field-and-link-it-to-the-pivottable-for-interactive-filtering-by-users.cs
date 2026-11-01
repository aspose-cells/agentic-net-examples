// Title: How to add a Region slicer linked to a PivotTable in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add slicer for PivotTable field Region | Create an interactive slicer for a PivotTable with Aspose.Cells .NET | Programmatically place a slicer at a specific cell using Aspose.Cells | Link a slicer to an existing PivotTable in C# with Aspose.Cells | Set slicer name and configure filtering for Excel workbook via Aspose.Cells
// Tags: Aspose.Cells add slicer to pivot table | C# slicer placement in Excel worksheet | link region field slicer to pivot table Aspose | programmatic Excel slicer creation .NET | pivot table interactive filtering with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // For PivotTable
using Aspose.Cells.Slicers; // For Slicer

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, checks for a PivotTable, and adds a slicer for the "Region" field at cell F6 (row 5, column 5) on the same worksheet. The slicer is named "RegionSlicer" and is linked to the first PivotTable, providing users with interactive filtering capabilities. The modified workbook is then saved to the specified output path.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook containing a PivotTable
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one PivotTable
                if (worksheet.PivotTables == null || worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No PivotTables found in the worksheet.");
                    return;
                }

                // Retrieve the first PivotTable
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Add a slicer for the "Region" field at cell F6 (row=5, column=5)
                // Note: The correct overload order is (PivotTable, firstRow, firstColumn, fieldName)
                int slicerIndex = worksheet.Slicers.Add(pivotTable, 5, 5, "Region");

                // Configure the slicer (optional: set a friendly name)
                Slicer slicer = worksheet.Slicers[slicerIndex];
                slicer.Name = "RegionSlicer";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the new slicer
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
