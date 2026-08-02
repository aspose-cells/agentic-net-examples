// Title: Load an Excel workbook without conditional formatting using Aspose.Cells for .NET
// Description: Demonstrates how to create a LoadFilter that excludes conditional formatting, assign it to LoadOptions, and open an XLSX file with Aspose.Cells. The resulting workbook loads faster and can be saved or processed without any conditional‑formatting rules.
// Keywords: Aspose.Cells LoadFilter | exclude conditional formatting | Excel loading performance .NET | LoadOptions conditional formatting | speed up workbook load
// Common Searches: Aspose.Cells skip conditional formatting when opening a file | how to improve Excel load speed with LoadFilter | disable conditional formatting on workbook load .NET | load large Excel workbook faster Aspose.Cells
// Developer Intent: Open an Excel file while omitting all conditional‑formatting rules to reduce load time and memory usage.
// Use Cases: Process massive analytics workbooks where visual styles are irrelevant, saving CPU and RAM. | Create lightweight copies of spreadsheets for downstream services that only need raw data. | Batch‑convert or export Excel files in a pipeline that must run quickly without rendering formatting.
// AI Prompts: Write C# code that loads an XLSX file with Aspose.Cells, disables conditional formatting, and saves the result. | Explain how to combine multiple LoadDataFilterOptions (e.g., charts, hyperlinks, conditional formatting) for selective loading. | Show a method to verify that no ConditionalFormattings exist after loading a workbook with a custom LoadFilter.

using System;
using Aspose.Cells;

// Demonstrates how to create a LoadFilter that excludes conditional formatting, assign it to LoadOptions, and open an XLSX file with Aspose.Cells. The resulting workbook loads faster and can be saved or processed without any conditional‑formatting rules.
class Program
{
    static void Main()
    {
        // Define filter options: load everything except conditional formatting
        LoadDataFilterOptions filterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.ConditionalFormatting;

        // Create a LoadFilter with the specified options
        LoadFilter loadFilter = new LoadFilter(filterOptions);

        // Set up LoadOptions and assign the custom filter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = loadFilter;

        // Load the workbook using the configured LoadOptions
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // (Optional) Verify that conditional formatting has been omitted
        // foreach (Worksheet sheet in workbook.Worksheets)
        // {
        //     Console.WriteLine($"Sheet '{sheet.Name}' conditional formatting count: {sheet.ConditionalFormattings.Count}");
        // }

        // Save the workbook after loading (if further processing or saving is required)
        workbook.Save("output.xlsx");
    }
}
