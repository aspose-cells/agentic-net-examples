// Title: C# – Load an XLSX workbook with Aspose.Cells while skipping all embedded charts
// Description: Demonstrates how to use Aspose.Cells LoadOptions and a LoadFilter to open an XLSX file from a path, clearing the Chart flag so no chart objects are loaded. The sample also shows how to verify the chart count after loading.
// Keywords: Aspose.Cells load workbook without charts | C# LoadOptions chart exclusion | LoadFilter SkipCharts Aspose.Cells | Open XLSX without chart data C# | Excel file performance Aspose.Cells
// Common Searches: how to ignore charts when loading Excel with Aspose.Cells | C# Aspose.Cells LoadOptions exclude chart objects | load XLSX file without charts for faster processing | verify chart count after opening workbook Aspose.Cells
// Developer Intent: Open an XLSX workbook in C# using Aspose.Cells while preventing any chart objects from being loaded into memory.
// Use Cases: Boost performance when processing large workbooks that contain many charts. | Extract only cell data for analytics, ignoring visual elements such as charts. | Validate that a workbook contains no charts after loading by checking each worksheet's Charts collection.
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells and excludes all charts using LoadOptions and LoadFilter. | Show how to load an XLSX workbook without charts and then save a chart‑free copy. | Explain how to combine LoadDataFilterOptions to skip multiple elements (e.g., charts, images) when loading a workbook with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells LoadOptions and a LoadFilter to open an XLSX file from a path, clearing the Chart flag so no chart objects are loaded. The sample also shows how to verify the chart count after loading.
class LoadWorkbookWithoutCharts
{
    static void Main()
    {
        // Path to the source XLSX file
        string filePath = "input.xlsx";

        // Create LoadOptions and set a LoadFilter that excludes charts
        LoadOptions loadOptions = new LoadOptions();
        // Load everything except charts (Chart flag = 256)
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart);

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Optional: verify that no charts were loaded
        int totalCharts = 0;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            totalCharts += sheet.Charts.Count;
        }
        Console.WriteLine($"Number of charts loaded: {totalCharts}");
    }
}
