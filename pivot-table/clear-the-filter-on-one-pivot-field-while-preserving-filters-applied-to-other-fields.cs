// Title: Clear a filter on a specific pivot field while preserving other pivot filters using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# clear filter on a single pivot field without affecting other fields | remove filter from pivot table field named Region using Aspose.Cells .NET | C# code to clear a specific pivot field filter while keeping other filters in an Excel workbook | example of Aspose.Cells ClearFilter method for a pivot field | how to programmatically clear a pivot field filter in Aspose.Cells for .NET
// Tags: Aspose.Cells ClearFilter pivot field | C# pivot table filter manipulation Aspose.Cells | remove specific pivot field filter .NET | preserve other pivot filters Aspose.Cells | access pivot field by name Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable and PivotField classes

// The example loads an Excel workbook, locates the first pivot table, searches for a pivot field named "Region" across row, column, and data fields, calls ClearFilter on that field to remove its filter while leaving all other field filters unchanged, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Access the first pivot table on the worksheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // Name of the pivot field whose filter should be cleared
            string fieldName = "Region";

            // Locate the pivot field (searching row, column and data fields)
            PivotField targetField = null;

            foreach (PivotField pf in pivotTable.RowFields)
            {
                if (pf.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                {
                    targetField = pf;
                    break;
                }
            }

            if (targetField == null)
            {
                foreach (PivotField pf in pivotTable.ColumnFields)
                {
                    if (pf.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetField = pf;
                        break;
                    }
                }
            }

            if (targetField == null)
            {
                foreach (PivotField pf in pivotTable.DataFields)
                {
                    if (pf.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetField = pf;
                        break;
                    }
                }
            }

            // Clear the filter on the identified field while preserving other filters
            if (targetField != null)
            {
                targetField.ClearFilter();
                Console.WriteLine($"Filter cleared for field \"{fieldName}\".");
            }
            else
            {
                Console.WriteLine($"Field \"{fieldName}\" not found in the pivot table.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
