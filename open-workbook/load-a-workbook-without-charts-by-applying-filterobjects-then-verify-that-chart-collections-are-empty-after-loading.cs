// Title: C# Example: Load an Excel workbook without charts using Aspose.Cells LoadFilter
// Description: Demonstrates how to create LoadOptions with a LoadFilter that excludes chart objects, open a workbook, verify that each worksheet's Charts collection is empty, and save a chart‑free copy. Ideal for improving load performance and preparing data‑only files.
// Keywords: Aspose.Cells LoadFilter exclude charts | C# load workbook without charts | LoadOptions chart filter .NET | verify empty chart collection Aspose.Cells | save chart‑free Excel file | Aspose.Cells performance optimization | Excel chart removal code sample
// Common Searches: Aspose.Cells load workbook without charts C# | How to filter out charts when opening Excel with Aspose.Cells | Check chart count after loading workbook Aspose.Cells | Create chart‑free copy of Excel file using Aspose.Cells | LoadOptions chart exclusion example
// Developer Intent: Open an Excel file while omitting all chart objects and confirm that no charts remain in any worksheet.
// Use Cases: Speed up processing of large reporting workbooks by skipping chart loading. | Generate data‑only versions of templates for downstream analytics or export. | Validate that a workbook intended for automated workflows contains no visual objects.
// AI Prompts: Write C# code that uses Aspose.Cells LoadFilter to open an Excel file without loading any charts and then saves the result. | Show how to combine LoadDataFilterOptions to exclude charts and verify the chart count per worksheet after loading. | Explain step‑by‑step how LoadOptions and LoadFilter work together to filter out chart objects in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to create LoadOptions with a LoadFilter that excludes chart objects, open a workbook, verify that each worksheet's Charts collection is empty, and save a chart‑free copy. Ideal for improving load performance and preparing data‑only files.
class Program
{
    static void Main()
    {
        // Path to the source workbook that contains charts
        string sourcePath = "ChartWorkbook.xlsx";

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Define filter options to load everything except charts
        LoadDataFilterOptions filterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart;

        // Initialize LoadFilter with the defined options
        LoadFilter loadFilter = new LoadFilter(filterOptions);

        // Assign the custom filter to LoadOptions
        loadOptions.LoadFilter = loadFilter;

        // Load the workbook using the filter that excludes charts
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Verify that chart collections are empty for each worksheet
        bool allChartsRemoved = true;
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws.Charts.Count > 0)
            {
                allChartsRemoved = false;
                Console.WriteLine($"Worksheet '{ws.Name}' still contains {ws.Charts.Count} chart(s).");
            }
        }

        if (allChartsRemoved)
        {
            Console.WriteLine("All charts were successfully filtered out during loading.");
        }

        // Save the workbook to confirm that no charts are present
        workbook.Save("ChartWorkbook_NoCharts.xlsx");
    }
}
