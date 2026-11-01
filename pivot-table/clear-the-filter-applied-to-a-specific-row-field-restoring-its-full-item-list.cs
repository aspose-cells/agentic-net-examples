// Title: How to clear a filter on a specific pivot table row field using Aspose.Cells for C#
// AI Prompts: Generate C# code with Aspose.Cells that locates a pivot row field by name and disables its filtering settings. | Show a full example that updates the pivot table after modifying the row field settings and writes the workbook to disk.
// Common Searches: Aspose.Cells C# remove filter from pivot table row field named Category | reset auto‑show and auto‑sort on a pivot field using Aspose.Cells | refresh pivot table after changing row field filter in .NET | example code to clear pivot row filter in a workbook with Aspose.Cells | how to programmatically unfilter a specific pivot row field in C#
// Tags: Aspose.Cells pivot row field filter removal | C# modify pivot table filter Aspose.Cells | disable IsAutoShow IsAutoSort pivot field | refresh pivot data after filter change Aspose.Cells | save workbook with updated pivot Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing workbook, accesses the first pivot table, finds the row field named "Category", turns off its IsAutoShow and IsAutoSort flags to remove the filter, refreshes and recalculates the pivot table, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the worksheet.");
                return;
            }

            // Access the first pivot table on the sheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Name of the row field whose filter should be cleared
            string targetRowFieldName = "Category"; // replace with your actual row field name

            // Locate the row field by name
            PivotField rowField = null;
            foreach (PivotField field in pivotTable.RowFields)
            {
                if (field.Name.Equals(targetRowFieldName, StringComparison.OrdinalIgnoreCase))
                {
                    rowField = field;
                    break;
                }
            }

            // If the row field is found, clear its filter
            if (rowField != null)
            {
                // Clear filter by disabling auto‑show and auto‑sort (Aspose.Cells does not expose a direct ClearFilter method in older versions)
                rowField.IsAutoShow = false;
                rowField.IsAutoSort = false;
            }
            else
            {
                Console.WriteLine($"Row field '{targetRowFieldName}' not found in the pivot table.");
            }

            // Refresh the pivot table to apply the change
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the cleared filter
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
