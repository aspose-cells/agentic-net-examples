// Title: C# – Load only “Summary” and “Data” worksheets with Aspose.Cells LoadFilter
// Description: Demonstrates how to create a custom LoadFilter that fully loads the "Summary" and "Data" sheets while loading only the structure for all other worksheets. The example shows configuring LoadOptions, opening the workbook, optionally removing unwanted sheets, and saving the filtered file, helping reduce memory usage and improve performance.
// Keywords: Aspose.Cells LoadFilter C# | load specific worksheets | partial workbook load | LoadOptions custom filter | memory‑efficient Excel processing | remove unwanted sheets Aspose | C# Excel sheet selection
// Common Searches: Aspose.Cells load only selected sheets C# | How to use LoadFilter to open specific worksheets | Load workbook with structure only for other sheets | C# remove worksheets after loading with Aspose.Cells | Partial Excel load for performance
// Developer Intent: Open an Excel file with Aspose.Cells, fully loading only the "Summary" and "Data" worksheets while keeping other sheets as structure‑only, then save a trimmed workbook.
// Use Cases: Speed up processing of large workbooks by loading only the needed sheets. | Create a lightweight copy containing just summary and data for reporting. | Pre‑filter worksheets before applying transformations to reduce memory footprint.
// AI Prompts: Generate a C# example that uses Aspose.Cells LoadFilter to load only worksheets named "Report" and "Details" and then saves the workbook. | Show how to modify CustomLoadFilter to load sheets matching a wildcard pattern like "*2023". | Explain how to handle formulas that reference sheets not loaded when using LoadFilter.

using System;
using Aspose.Cells;

namespace LoadSpecificSheetsExample
{
    // Custom LoadFilter to load data only for "Summary" and "Data" worksheets
    // Demonstrates how to create a custom LoadFilter that fully loads the "Summary" and "Data" sheets while loading only the structure for all other worksheets. The example shows configuring LoadOptions, opening the workbook, optionally removing unwanted sheets, and saving the filtered file, helping reduce memory usage and improve performance.
    class CustomLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load full data for the required sheets, otherwise load only the structure
            if (sheet.Name == "Summary" || sheet.Name == "Data")
            {
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            }
            else
            {
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Configure LoadOptions with the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new CustomLoadFilter();

            // Load the workbook using the specified LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Optional: remove worksheets that are not needed after loading
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                Worksheet ws = workbook.Worksheets[i];
                if (ws.Name != "Summary" && ws.Name != "Data")
                {
                    workbook.Worksheets.RemoveAt(i);
                }
            }

            // Save the filtered workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}
