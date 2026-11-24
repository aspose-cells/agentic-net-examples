// Title: Add a Region slicer to a PivotTable programmatically with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells code to add a slicer for the Region field in a PivotTable | how to link an Excel slicer to a PivotTable using Aspose.Cells .NET | example of inserting a slicer at a specific cell location with Aspose.Cells
// Tags: Aspose.Cells add slicer to PivotTable | C# programmatic Excel slicer insertion | associate slicer with PivotTable field | apply slicer visual style Aspose.Cells | save workbook after slicer modification

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // For PivotTable related classes
using Aspose.Cells.Slicers; // For Slicer related classes

namespace AsposeCellsExample
{
    // The example loads an existing workbook, checks for a PivotTable on the first worksheet, adds a slicer for the "Region" field at row 5, column 5 linked to that PivotTable, optionally sets a visual style, and saves the updated file.
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
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one PivotTable
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No PivotTables found on the first worksheet.");
                    return;
                }

                // Retrieve the first PivotTable
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Add a slicer for the "Region" field at row 5, column 5
                // Parameter order: (pivotTable, firstRow, firstColumn, fieldName)
                int slicerIndex = worksheet.Slicers.Add(pivotTable, 5, 5, "Region");
                Slicer slicer = worksheet.Slicers[slicerIndex];

                // Optional: set a visual style for the slicer if supported
                // Uncomment the following line if your Aspose.Cells version provides SlicerStyle
                // slicer.SlicerStyle = SlicerStyleType.SlicerStyleLight1;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
