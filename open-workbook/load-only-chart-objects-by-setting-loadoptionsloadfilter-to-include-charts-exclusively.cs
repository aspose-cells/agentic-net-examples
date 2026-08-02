// Title: Load Only Chart Objects from an Excel Workbook with Aspose.Cells C#
// Description: Demonstrates how to use Aspose.Cells LoadOptions with a LoadFilter set to LoadDataFilterOptions.Chart to open an Excel file, load exclusively its chart objects, enumerate them, and save a new workbook that contains only the charts, reducing memory usage and processing time.
// Keywords: Aspose.Cells C# LoadFilter | LoadDataFilterOptions.Chart | extract charts from Excel | chart‑only workbook | performance‑optimized Excel loading | save workbook with charts only | Aspose.Cells LoadOptions example
// Common Searches: Aspose.Cells load only charts C# | LoadFilter chart objects Excel | How to save chart‑only workbook with Aspose | C# extract charts from .xlsx using Aspose.Cells | Load workbook without cell data Aspose
// Developer Intent: Open an Excel file, keep just the chart objects, and write a new file that contains only those charts.
// Use Cases: Create a lightweight version of a report that includes only visual charts for quick distribution. | Separate chart graphics from a data‑heavy workbook for presentation slides. | Count or validate charts per worksheet without incurring the overhead of loading cell data.
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, using LoadOptions to include only chart objects, then saves the result. | Explain how to enumerate worksheets and retrieve the chart count after applying LoadDataFilterOptions.Chart. | Show how to produce a chart‑only workbook while skipping all cell data for faster processing.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells LoadOptions with a LoadFilter set to LoadDataFilterOptions.Chart to open an Excel file, load exclusively its chart objects, enumerate them, and save a new workbook that contains only the charts, reducing memory usage and processing time.
class LoadChartsOnly
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Path to the output workbook (will contain only charts)
        string outputPath = "output_charts_only.xlsx";

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Configure the LoadFilter to load only chart objects
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.Chart);

        // Load the workbook with the specified load options
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Optional: Verify that only charts are loaded
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Sheet: {sheet.Name}, Charts loaded: {sheet.Charts.Count}");
        }

        // Save the workbook; it will contain only the chart objects
        workbook.Save(outputPath);
    }
}
