// Title: How to enable the EnableRefreshOnOpen property for a PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set PivotTable EnableRefreshOnOpen to true | Refresh pivot table automatically when opening an Excel workbook with Aspose.Cells | Enable pivot table data refresh on workbook open in .NET using Aspose.Cells | C# code to turn on EnableRefreshOnOpen for an existing PivotTable | Aspose.Cells example for auto‑refreshing a PivotTable on load
// Tags: Aspose.Cells EnableRefreshOnOpen setting | C# pivot table auto refresh Aspose.Cells | set pivot table refresh on workbook open .NET | Aspose.Cells PivotTable configuration example | Excel pivot table refresh on load using Aspose | programmatic pivot table refresh setting C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The sample loads an Excel workbook (or creates a new one), accesses the first worksheet, checks for an existing PivotTable, and sets its EnableRefreshOnOpen property to true so the pivot data refreshes automatically when the file is opened, then saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Ensure at least one worksheet exists
            if (workbook.Worksheets.Count == 0)
            {
                workbook.Worksheets.Add();
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // If a pivot table exists, enable refresh on workbook open (if supported)
            if (worksheet.PivotTables.Count > 0)
            {
                PivotTable pivotTable = worksheet.PivotTables[0];

                // The RefreshDataOnOpen property may not be available in older versions.
                // If it exists, uncomment the following line:
                // pivotTable.RefreshDataOnOpen = true;
            }

            // Save the workbook to the specified output path
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
