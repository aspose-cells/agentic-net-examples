// Title: Remove all custom groupings from PivotTable base fields using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# snippet with Aspose.Cells that iterates over every base field of the first pivot table, calls Ungroup on each field, and saves the workbook. | Show how to create a reusable method that accepts an input Excel file path, clears all custom groups from its pivot table fields using Aspose.Cells, and writes the result to an output file. | Demonstrate the steps to verify that a worksheet contains a pivot table before performing Ungroup operations on its fields in Aspose.Cells.
// Common Searches: C# Aspose.Cells clear grouping on pivot table base fields | how to ungroup all fields in an Excel pivot table using Aspose.Cells .NET | programmatically delete custom groups from pivot fields with Aspose.Cells | remove custom grouping from pivot table base fields Aspose.Cells C#
// Tags: Aspose.Cells clear pivot field groups | C# ungroup pivot table fields Aspose.Cells | remove custom grouping from pivot base fields .NET | iterate pivot table base fields Aspose.Cells | save modified workbook Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, checks for a pivot table on the first worksheet, loops through each base field of that pivot table, calls Ungroup to delete any custom grouping, and then saves the updated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the workbook that contains the pivot table
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one pivot table in the worksheet
        if (worksheet.PivotTables.Count > 0)
        {
            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Iterate through all base fields of the pivot table
            // BaseFields includes both row and column fields that can be grouped
            foreach (PivotField field in pivotTable.BaseFields)
            {
                // Remove any custom grouping applied to the field
                // The Ungroup method clears all grouping for the field
                field.Ungroup();
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
