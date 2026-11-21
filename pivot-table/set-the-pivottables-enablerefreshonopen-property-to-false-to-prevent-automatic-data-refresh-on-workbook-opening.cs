// Title: How to disable a PivotTable’s automatic refresh on workbook open with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set PivotTable EnableRefreshOnOpen to false | prevent Excel pivot table from refreshing when opening file using Aspose.Cells | disable automatic pivot refresh on workbook open in .NET | how to turn off RefreshDataOnOpen for a PivotTable with Aspose.Cells | C# code to stop pivot table data refresh on opening workbook Aspose.Cells | Aspose.Cells pivot table refresh setting not working on open
// Tags: Aspose.Cells PivotTable DisableRefreshOnOpen | C# set PivotTable EnableRefreshOnOpen false | Aspose.Cells prevent pivot refresh on open | Excel workbook pivot refresh control .NET | Aspose.Cells reflection set pivot property

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, checks for a PivotTable on the first worksheet, and uses reflection to set the PivotTable's EnableRefreshOnOpen (or fallback RefreshDataOnOpen) property to false, preventing automatic data refresh when the workbook is opened, then saves the modified file.
class Program
{
    static void Main()
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

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one PivotTable
            if (sheet.PivotTables == null || sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                workbook.Save(outputPath);
                return;
            }

            // Get the first PivotTable
            PivotTable pivotTable = sheet.PivotTables[0];

            // Attempt to disable automatic refresh on open (property may vary by version)
            try
            {
                var prop = typeof(PivotTable).GetProperty("EnableRefreshOnOpen");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(pivotTable, false);
                }
                else
                {
                    // Fallback to an alternative property name if present
                    var altProp = typeof(PivotTable).GetProperty("RefreshDataOnOpen");
                    if (altProp != null && altProp.CanWrite)
                    {
                        altProp.SetValue(pivotTable, false);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to set refresh property: {ex.Message}");
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
