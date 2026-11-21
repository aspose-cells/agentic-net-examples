// Title: Remove all filters from an Excel pivot table using Aspose.Cells for .NET (C#)
// Common Searches: C# code to clear every filter on an Excel pivot table with Aspose.Cells | Aspose.Cells .NET how to reset pivot table filters programmatically | Remove row, column, page, and data field filters from a pivot table in C# | Clear pivot table filters and display all data using Aspose.Cells workbook | Aspose.Cells ClearFilters method example for pivot tables in .NET
// Tags: Aspose.Cells ClearFilters API | C# unfilter Excel pivot table | Aspose.Cells pivot table filter reset | programmatic Excel pivot table filter removal .NET | Aspose.Cells workbook pivot table filter clearing

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an Excel workbook, checks for a pivot table on the first worksheet, calls ClearFilters() on the pivot table, clears filters on all row, column, page, and data fields, and saves the workbook so the pivot table displays unfiltered data.
class RemovePivotFilters
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Get the first pivot table on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Clear all filters applied to the pivot table
            pivotTable.ClearFilters();

            // Additionally, clear filters on each field type to ensure no residual filters remain
            foreach (PivotField field in pivotTable.RowFields)
            {
                field.ClearFilter();
            }

            foreach (PivotField field in pivotTable.ColumnFields)
            {
                field.ClearFilter();
            }

            foreach (PivotField field in pivotTable.PageFields)
            {
                field.ClearFilter();
            }

            foreach (PivotField field in pivotTable.DataFields)
            {
                field.ClearFilter();
            }

            // Save the workbook with the pivot table now showing unfiltered data
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
