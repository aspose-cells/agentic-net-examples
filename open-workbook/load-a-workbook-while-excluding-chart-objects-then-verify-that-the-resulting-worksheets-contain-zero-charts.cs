// Title: C# – Load an Excel workbook without charts using Aspose.Cells LoadFilter and verify chart count
// Description: Demonstrates how to create a custom LoadFilter that disables chart loading, open an XLSX file with LoadOptions, iterate through each worksheet to confirm the chart collection is empty, and properly dispose the workbook.
// Keywords: Aspose.Cells LoadFilter exclude charts | C# load workbook without charts | LoadOptions chart exclusion | Verify worksheet chart count Aspose | Excel performance chart removal | LoadDataFilterOptions Chart | Aspose.Cells custom LoadFilter example
// Common Searches: how to load excel file without charts using aspose.cells | asp.net exclude charts when loading workbook | c# check worksheet chart count after loading | aspocells loadoptions chart filter | remove chart objects during workbook load
// Developer Intent: Load an Excel workbook while omitting all chart objects and confirm that each worksheet contains zero charts.
// Use Cases: Speed up processing of large workbooks when charts are irrelevant. | Validate that a template file is chart‑free before applying data updates. | Prepare workbooks for server‑side calculations without rendering overhead.
// AI Prompts: Show me a C# example of a custom LoadFilter in Aspose.Cells that excludes charts and prints the chart count for each worksheet. | Give an alternative method to strip charts from a workbook after it has been loaded with Aspose.Cells. | Explain how LoadDataFilterOptions can be combined to load only data and skip charts in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a custom LoadFilter that disables chart loading, open an XLSX file with LoadOptions, iterate through each worksheet to confirm the chart collection is empty, and properly dispose the workbook.
class Program
{
    static void Main()
    {
        // Create LoadOptions and assign a custom LoadFilter that excludes charts
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new ExcludeChartLoadFilter();

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Verify that each worksheet contains zero charts
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet '{sheet.Name}' chart count: {sheet.Charts.Count}");
            if (sheet.Charts.Count != 0)
            {
                Console.WriteLine("Warning: Charts were not excluded from this worksheet.");
            }
        }

        // Clean up
        workbook.Dispose();
    }

    // Custom LoadFilter implementation to exclude chart objects during loading
    private class ExcludeChartLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load all data except charts
            this.LoadDataFilterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart;
        }
    }
}
