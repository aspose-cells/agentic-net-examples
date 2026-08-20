// Title: Load an Excel workbook without charts using Aspose.Cells LoadFilter (C#)
// Description: Demonstrates how to create a custom LoadFilter that disables the LoadDataFilterOptions.Chart flag, load a workbook with charts omitted, verify that each worksheet's chart collection is empty, and optionally save the chart‑free file.
// Keywords: Aspose.Cells LoadFilter C# | exclude charts when loading Excel | LoadDataFilterOptions Chart false | verify empty chart collection | skip chart objects Aspose.Cells
// Common Searches: Aspose.Cells load workbook without charts | C# load Excel file excluding charts | How to check chart count after loading with Aspose.Cells | Custom LoadFilter to omit charts in .NET
// Developer Intent: Open an Excel file while skipping all chart objects and confirm that no charts were loaded.
// Use Cases: Boost performance when processing large workbooks that contain many charts but only data is needed. | Ensure a source workbook is chart‑free before data extraction or transformation. | Create a lightweight copy of a workbook for distribution or archival without visual elements.
// AI Prompts: Write C# code using Aspose.Cells to load a workbook without charts and output the chart count per worksheet. | Explain how to combine LoadDataFilterOptions flags to exclude specific object types such as charts when opening an Excel file. | Show how to extend the custom LoadFilter to also skip images, shapes, or other drawing objects.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a custom LoadFilter that disables the LoadDataFilterOptions.Chart flag, load a workbook with charts omitted, verify that each worksheet's chart collection is empty, and optionally save the chart‑free file.
class NoChartLoadFilter : LoadFilter
{
    // Override StartSheet to set filter options that exclude charts
    public override void StartSheet(Worksheet sheet)
    {
        // Load all data except charts
        this.LoadDataFilterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart;
    }
}

class Program
{
    static void Main()
    {
        // Path to the source workbook that contains charts
        string sourcePath = "input.xlsx";

        // Create LoadOptions and assign the custom filter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new NoChartLoadFilter();

        // Load the workbook using the filter (charts will not be loaded)
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Verify that each worksheet has an empty chart collection
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet '{sheet.Name}' chart count: {sheet.Charts.Count}");
        }

        // Save the workbook (optional, demonstrates that charts are absent)
        workbook.Save("output.xlsx");
    }
}
