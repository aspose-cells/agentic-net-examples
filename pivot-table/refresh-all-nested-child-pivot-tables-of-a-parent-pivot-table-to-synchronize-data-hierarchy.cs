// Title: Refresh all child pivot tables of a parent pivot table in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells refresh child pivot tables after data change | How to programmatically refresh nested pivot tables with Aspose.Cells | Refresh every pivot table in an Excel workbook using Aspose.Cells .NET | Iterate worksheets and call RefreshData on each pivot table Aspose.Cells example | Update parent‑child pivot table hierarchy with Aspose.Cells C# code
// Tags: Aspose.Cells RefreshData pivot tables | C# refresh nested pivot tables Excel | Iterate pivot tables workbook Aspose.Cells | Parent child pivot table synchronization Aspose.Cells | Refresh all pivot tables in Excel using .NET

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, finds a specified parent pivot table, then loops through all worksheets and their pivot tables, invoking RefreshData on each to synchronize caches for parent and child pivots, and finally saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains the pivot tables
            Workbook workbook = new Workbook(inputPath);

            // Name of the parent pivot table whose child pivots need to be refreshed
            const string parentPivotName = "ParentPivot";

            // Locate the parent pivot table in the workbook
            PivotTable parentPivot = null;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pt in sheet.PivotTables)
                {
                    if (pt.Name == parentPivotName)
                    {
                        parentPivot = pt;
                        break;
                    }
                }
                if (parentPivot != null)
                    break;
            }

            if (parentPivot == null)
            {
                Console.WriteLine($"Parent pivot table \"{parentPivotName}\" not found.");
                return;
            }

            // Refresh all pivot tables (including any child pivots) in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pt in sheet.PivotTables)
                {
                    try
                    {
                        pt.RefreshData(); // Refreshes the pivot cache and data
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to refresh pivot table \"{pt.Name}\": {ex.Message}");
                    }
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
