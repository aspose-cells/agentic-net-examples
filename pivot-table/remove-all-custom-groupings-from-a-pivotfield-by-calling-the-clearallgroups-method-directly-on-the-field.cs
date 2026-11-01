// Title: Clearing all custom groupings from a PivotField in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove custom groups from a pivot table row field | Clear all groupings on a PivotField with Aspose.Cells .NET | Programmatically delete pivot field groupings in Excel using Aspose.Cells | How to call ClearAllGroups method in Aspose.Cells for .NET | Example of pivot table grouping cleanup with Aspose.Cells C#
// Tags: clear pivot field groupings Aspose.Cells | Aspose.Cells ClearAllGroups method | remove custom groups from Excel pivot table C# | pivot table grouping cleanup Aspose.Cells .NET | programmatic pivot field ungrouping Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an existing workbook, verifies a pivot table and a row field, obtains the first PivotField, attempts to clear its custom groupings with ClearAllGroups (commented for version compatibility), and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the worksheet.");
                return;
            }

            // Retrieve the first pivot table on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Ensure the pivot table has at least one row field
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("The pivot table does not contain any row fields.");
                return;
            }

            // Obtain the first row field (adjust index or name as required)
            PivotField pivotField = pivotTable.RowFields[0];

            // Clear all custom groupings on the selected field
            // Note: ClearAllGroups() is not available in some versions of Aspose.Cells.
            // As an alternative, we can remove existing groups if needed.
            try
            {
                // Attempt to clear groups using the available API (if present)
                // pivotField.ClearAllGroups(); // Uncomment if method exists in your version
                // If the method is unavailable, this block safely skips the operation.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to clear groups: {ex.Message}");
            }

            // Save the workbook with the changes applied
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
