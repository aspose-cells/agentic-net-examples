// Title: C# – Load Only Chart Objects from an Excel Workbook Using Aspose.Cells LoadFilter
// Description: Demonstrates how to configure Aspose.Cells LoadOptions with a LoadFilter set to LoadDataFilterOptions.Chart, load an XLSX file so that only chart objects are retained, enumerate the charts per worksheet, and save a new workbook containing just the charts.
// Keywords: Aspose.Cells LoadFilter chart | LoadDataFilterOptions.Chart C# | load workbook charts only | extract charts Aspose.Cells | save workbook with charts only | .NET Excel chart extraction
// Common Searches: Aspose.Cells load only charts example | C# LoadFilter to include just chart objects | how to extract charts from Excel with Aspose.Cells | save Excel file with only charts using .NET | LoadOptions chart-only loading Aspose
// Developer Intent: Load an Excel file, keep only its chart objects, and write them to a new workbook.
// Use Cases: Create a lightweight workbook that contains only visualizations for distribution. | Extract charts from a data‑heavy report to generate a summary file for mobile viewers. | Validate chart presence and count without loading full worksheet data, reducing memory usage.
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions with LoadFilter = LoadDataFilterOptions.Chart to load only charts from an Excel file and save them to a new workbook. | Explain the impact of LoadFilter on workbook loading in Aspose.Cells and show how to list the charts after applying the filter. | Suggest performance‑optimized practices for extracting chart objects from large workbooks with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to configure Aspose.Cells LoadOptions with a LoadFilter set to LoadDataFilterOptions.Chart, load an XLSX file so that only chart objects are retained, enumerate the charts per worksheet, and save a new workbook containing just the charts.
class LoadChartsOnly
{
    static void Main()
    {
        // Input workbook containing various data and charts
        string inputFile = "input.xlsx";

        // Output workbook that will contain only the chart objects
        string outputFile = "charts_only.xlsx";

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Configure the LoadFilter to load only charts
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.Chart);

        // Load the workbook with the specified load options
        Workbook workbook = new Workbook(inputFile, loadOptions);

        // Verify that only charts are loaded (optional)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Sheet: {sheet.Name}, Charts loaded: {sheet.Charts.Count}");
        }

        // Save the workbook; it will contain only the chart objects
        workbook.Save(outputFile);
    }
}
