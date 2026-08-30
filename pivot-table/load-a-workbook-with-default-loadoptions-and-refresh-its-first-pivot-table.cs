// Title: Refresh the first pivot table in an Excel workbook using Aspose.Cells for .NET with default LoadOptions
// AI Prompts: Load an .xlsx file with Aspose.Cells default LoadOptions, locate the initial pivot table on the first worksheet, call RefreshData and CalculateData, and then save the workbook. | Generate C# code that checks for a pivot table on the first sheet, refreshes its source data, recalculates the pivot, and writes the result to a new file using Aspose.Cells.
// Common Searches: Aspose.Cells C# refresh first pivot table after loading workbook | how to use LoadOptions to open workbook and update pivot table with Aspose.Cells | programmatically refresh pivot table data in Excel using Aspose.Cells .NET | C# example for RefreshData and CalculateData on a pivot table with Aspose.Cells | saving workbook after pivot table refresh using Aspose.Cells for .NET
// Tags: Aspose.Cells refresh pivot table C# | default LoadOptions workbook load Aspose.Cells | pivot table RefreshData method .NET | pivot table CalculateData operation Aspose.Cells | persist workbook after pivot refresh Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The program loads 'input.xlsx' with default LoadOptions, accesses the first worksheet, verifies a pivot table exists, refreshes its data via RefreshData(), recalculates it with CalculateData(), and saves the modified workbook as 'output.xlsx'.
class RefreshFirstPivotTable
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Load the workbook with default LoadOptions
        LoadOptions loadOptions = new LoadOptions();               // default options
        Workbook workbook = new Workbook(inputPath, loadOptions); // load with options

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one pivot table
        if (worksheet.PivotTables.Count > 0)
        {
            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Refresh the pivot table data
            pivotTable.RefreshData();

            // Optionally recalculate the pivot table after refresh
            pivotTable.CalculateData();
        }
        else
        {
            Console.WriteLine("No pivot tables found in the first worksheet.");
        }

        // Save the workbook (optional)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
