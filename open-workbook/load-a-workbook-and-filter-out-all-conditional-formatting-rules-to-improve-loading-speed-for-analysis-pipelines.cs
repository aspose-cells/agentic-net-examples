// Title: C# – Load an Excel Workbook Excluding Conditional Formatting with Aspose.Cells
// Description: Shows how to configure Aspose.Cells LoadFilter and LoadOptions to open an .xlsx file while skipping all conditional‑formatting rules, cutting memory consumption and load time. After loading, the workbook can be processed or saved.
// Keywords: Aspose.Cells | C# | .NET | LoadFilter | LoadOptions | ConditionalFormatting | LoadDataFilterOptions | excel load performance | skip conditional formatting | large workbook processing | data pipeline
// Common Searches: Aspose.Cells load workbook without conditional formatting | C# skip conditional formatting when opening Excel file | How to improve Excel load speed with Aspose.Cells | LoadFilter example for conditional formatting | Reduce memory usage when loading large Excel files .NET
// Developer Intent: The developer wants to open an Excel workbook quickly by omitting conditional‑formatting data to improve performance in analysis pipelines.
// Use Cases: Processing massive spreadsheets where visual styles are irrelevant | Running automated data extraction or transformation jobs with minimal overhead | Generating raw data reports from many workbooks in a CI/CD pipeline | Performing statistical analysis on Excel files without the cost of rendering formatting
// AI Prompts: Provide a C# example that loads an Excel file with Aspose.Cells while excluding conditional formatting. | Explain how LoadDataFilterOptions.All & ~LoadDataFilterOptions.ConditionalFormatting works in Aspose.Cells. | Show how to configure LoadOptions to improve workbook load speed for large files. | Give a step‑by‑step guide to skip conditional formatting when opening a workbook and then save it.

using System;
using Aspose.Cells;

// Shows how to configure Aspose.Cells LoadFilter and LoadOptions to open an .xlsx file while skipping all conditional‑formatting rules, cutting memory consumption and load time. After loading, the workbook can be processed or saved.
class Program
{
    static void Main()
    {
        // Create a LoadFilter that loads everything except conditional formatting
        LoadFilter filter = new LoadFilter();
        // Exclude the ConditionalFormatting flag from the default "All" options
        filter.LoadDataFilterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.ConditionalFormatting;

        // Set the filter in LoadOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = filter;

        // Load the workbook using the configured LoadOptions
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // (Optional) Save the workbook after loading
        workbook.Save("output.xlsx");
    }
}
