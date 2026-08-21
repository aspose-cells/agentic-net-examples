// Title: C# Example: Load Workbook with Default LoadOptions and Refresh the First Pivot Table using Aspose.Cells
// Description: Demonstrates how to load an Excel workbook with Aspose.Cells' default LoadOptions, detect the first worksheet's pivot tables, refresh the data source, recalculate the pivot, and optionally save the updated file.
// Keywords: Aspose.Cells | C# pivot table refresh | LoadOptions | refresh first pivot table | calculate pivot data | Excel workbook load | programmatic pivot update | default LoadOptions | pivot table API | Aspose.Cells example
// Common Searches: Aspose.Cells refresh pivot table C# | load Excel workbook with default options Aspose | how to refresh first pivot table using Aspose.Cells | C# code to refresh pivot tables after loading workbook | Aspose.Cells example refresh pivot data
// Developer Intent: Refresh the first pivot table after loading a workbook with default options.
// Use Cases: Update pivot data after external source changes before saving the workbook. | Automate pivot refresh in batch processing of multiple Excel files. | Validate pivot calculations during data transformation pipelines. | Integrate pivot refresh into server‑side reporting services.
// AI Prompts: Write C# code that loads an Excel file with default LoadOptions and refreshes every pivot table in all worksheets using Aspose.Cells. | Show how to handle worksheets that contain no pivot tables while performing a refresh with Aspose.Cells. | Explain exception handling for PivotTable.RefreshData and CalculateData methods in Aspose.Cells. | Provide a GitHub‑style README snippet describing this example, its prerequisites, and how to run it.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to load an Excel workbook with Aspose.Cells' default LoadOptions, detect the first worksheet's pivot tables, refresh the data source, recalculate the pivot, and optionally save the updated file.
class Program
{
    static void Main()
    {
        // Path to the workbook to be loaded
        string inputPath = "input.xlsx";

        // Load the workbook with default LoadOptions
        LoadOptions loadOptions = new LoadOptions();               // default options
        Workbook workbook = new Workbook(inputPath, loadOptions); // load with options

        // Refresh the first pivot table if it exists
        if (workbook.Worksheets[0].PivotTables.Count > 0)
        {
            PivotTable firstPivot = workbook.Worksheets[0].PivotTables[0];
            firstPivot.RefreshData();      // refresh data from the source
            firstPivot.CalculateData();    // recalculate the pivot after refresh (optional)
        }

        // Save the updated workbook (optional)
        workbook.Save("output.xlsx");
    }
}
