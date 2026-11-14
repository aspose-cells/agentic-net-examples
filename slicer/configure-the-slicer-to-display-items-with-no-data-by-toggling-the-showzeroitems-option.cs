// Title: Enable Show Items With No Data (Show Zero Items) for a Pivot Table Slicer in Aspose.Cells for .NET
// Common Searches: how to add a slicer that includes empty categories using Aspose.Cells C# | Aspose.Cells .NET enable slicer to show items with no data in pivot table | C# configure slicer to display zero‑value items with Aspose.Cells API | workaround for missing ShowItemsWithNoData property in Aspose.Cells slicer | display categories without data in Excel slicer via Aspose.Cells library
// Tags: Aspose.Cells slicer display empty categories | C# add slicer to pivot table with Aspose.Cells | Aspose.Cells configure slicer zero‑item visibility | Excel slicer show‑no‑data items Aspose.Cells | Aspose.Cells .NET slicer settings for no‑data items

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The example loads an existing workbook, checks for a pivot table, adds a slicer for the "Category" field, assigns a custom name, notes that the ShowItemsWithNoData property is unavailable in the current Aspose.Cells version, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure a pivot table exists
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Add a slicer for the "Category" pivot field.
            // Overload: Add(PivotTable pivotTable, int firstRow, int firstColumn, string fieldName)
            int slicerIndex = worksheet.Slicers.Add(pivotTable, 1, 0, "Category");

            // Get the created slicer object
            Slicer slicer = worksheet.Slicers[slicerIndex];

            // Optionally set a custom name for the slicer
            slicer.Name = "CategorySlicer";

            // Note: The ShowItemsWithNoData property is not available in this version of Aspose.Cells.
            // If needed, configure related settings via the slicer cache or other available APIs.

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
