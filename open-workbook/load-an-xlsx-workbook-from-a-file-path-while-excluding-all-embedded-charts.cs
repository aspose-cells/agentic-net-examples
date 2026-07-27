// Title: C# – Load an XLSX Workbook Without Charts Using Aspose.Cells LoadOptions
// Description: Demonstrates how to open an XLSX file with Aspose.Cells for .NET while skipping every embedded chart. The example creates a LoadOptions object, sets a LoadFilter that excludes charts via LoadDataFilterOptions, loads the workbook from a file path, and verifies that no chart objects are present.
// Keywords: Aspose.Cells load workbook without charts | C# LoadOptions chart exclusion | LoadDataFilterOptions chart | Aspose.Cells LoadFilter example | skip charts when loading Excel | performance Excel loading Aspose | server‑side Excel processing no charts | Aspose.Cells .NET chart filter | exclude embedded charts Aspose | load XLSX file without chart objects
// Common Searches: Aspose.Cells load XLSX without charts C# | How to exclude charts when opening Excel with Aspose.Cells | LoadFilter to skip charts in Aspose.Cells .NET | C# open workbook without chart objects | Improve Excel load performance by omitting charts Aspose
// Developer Intent: Open an XLSX workbook from a file path while preventing any chart objects from being loaded into memory.
// Use Cases: Speed up processing of large workbooks that contain many charts by loading only data and formulas. | Create a data‑only copy of a workbook for server‑side calculations, analytics, or validation without the overhead of chart rendering. | Perform worksheet transformations or data extraction while ensuring charts are excluded from the in‑memory model.
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells and excludes all charts using LoadOptions and LoadFilter. | Explain how LoadDataFilterOptions can be combined to load specific workbook parts while omitting charts in Aspose.Cells. | Provide a step‑by‑step guide to verify that no charts were loaded after opening a workbook with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to open an XLSX file with Aspose.Cells for .NET while skipping every embedded chart. The example creates a LoadOptions object, sets a LoadFilter that excludes charts via LoadDataFilterOptions, loads the workbook from a file path, and verifies that no chart objects are present.
class Program
{
    static void Main()
    {
        // Path to the XLSX file to be loaded
        string filePath = "input.xlsx";

        // Create LoadOptions
        LoadOptions loadOptions = new LoadOptions();

        // Define filter options that load everything except charts
        LoadDataFilterOptions filterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart;

        // Assign a LoadFilter with the defined options
        loadOptions.LoadFilter = new LoadFilter(filterOptions);

        // Load the workbook using the file path and the configured LoadOptions
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Verify that no charts were loaded
        int totalCharts = 0;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            totalCharts += sheet.Charts.Count;
        }
        Console.WriteLine($"Charts loaded: {totalCharts}");
    }
}
