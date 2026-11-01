// Title: Refresh parent pivot table and recalculate all child pivot tables in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# refresh parent pivot table and recalculate child pivots | How to update totals of multiple pivot tables after data refresh with Aspose.Cells | C# code to refresh and calculate all pivot tables in an Excel file using Aspose.Cells | Iterate through worksheets to refresh child pivot tables after parent pivot refresh Aspose.Cells | RefreshData and CalculateData for each pivot table in Aspose.Cells .NET
// Tags: Aspose.Cells refreshdata calculatedata API | process multiple pivot tables .NET | iterate worksheets for pivot tables Aspose.Cells | update pivot table totals Excel C# | parent-child pivot table synchronization Aspose.Cells | refresh all pivot tables workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, identifies the first pivot table as the parent, refreshes its source data and recalculates its totals, then iterates through every other pivot table across all worksheets, refreshing and recalculating each child pivot table before saving the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "ParentWorkbook.xlsx";
        const string outputPath = "ParentWorkbook_Updated.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Locate the first pivot table (treated as the parent)
            PivotTable parentPivot = null;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.PivotTables.Count > 0)
                {
                    parentPivot = sheet.PivotTables[0];
                    break;
                }
            }

            // Refresh and calculate the parent pivot table if found
            if (parentPivot != null)
            {
                parentPivot.RefreshData();   // Refresh source data
                parentPivot.CalculateData(); // Recalculate totals and subtotals
            }

            // Process all other pivot tables (children)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    if (pivot == parentPivot)
                        continue; // Skip already processed parent

                    pivot.RefreshData();
                    pivot.CalculateData();
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
