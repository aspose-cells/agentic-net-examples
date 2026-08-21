// Title: Load only "Summary" and "Data" worksheets with Aspose.Cells LoadFilter in C#
// Description: Demonstrates a custom LoadFilter that fully loads the "Summary" and "Data" sheets (LoadDataFilterOptions.All) and loads only the structure for all other worksheets (LoadDataFilterOptions.Structure). The filter is applied via LoadOptions, the workbook is opened, sheet cell counts are displayed, and the file is saved.
// Keywords: Aspose.Cells | LoadFilter | C# | load specific worksheets | partial workbook loading | LoadDataFilterOptions.All | LoadDataFilterOptions.Structure | memory optimization | Excel sheet selection | custom load filter
// Common Searches: Aspose.Cells load only selected sheets | C# LoadFilter example for specific worksheets | partial workbook loading with LoadOptions | how to skip sheet data in Aspose.Cells | load workbook structure only for some sheets
// Developer Intent: Open a workbook, fully loading only the "Summary" and "Data" worksheets while loading just the layout of all other sheets.
// Use Cases: Minimize RAM usage when processing large workbooks by loading full data only for required sheets. | Create a lightweight copy of a file that retains complete data for essential sheets and only the layout for ancillary ones. | Generate reports that need detailed values from "Summary" and "Data" while ignoring cell contents of other worksheets.
// AI Prompts: Write a C# snippet that uses Aspose.Cells LoadOptions with a custom LoadFilter to fully load "Summary" and "Data" sheets and load only the structure for the rest. | Explain the performance impact of LoadDataFilterOptions.All versus LoadDataFilterOptions.Structure when opening large Excel files with Aspose.Cells. | Show an alternative way to load selected worksheets without a custom LoadFilter, using built‑in Aspose.Cells features.

using System;
using Aspose.Cells;

namespace LoadSpecificSheetsExample
{
    // Custom load filter to load only "Summary" and "Data" worksheets with full data.
    // Other worksheets will be loaded with only their structure (no cell data).
    // Demonstrates a custom LoadFilter that fully loads the "Summary" and "Data" sheets (LoadDataFilterOptions.All) and loads only the structure for all other worksheets (LoadDataFilterOptions.Structure). The filter is applied via LoadOptions, the workbook is opened, sheet cell counts are displayed, and the file is saved.
    public class CustomLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Check the worksheet name and set the appropriate load options.
            if (sheet.Name == "Summary" || sheet.Name == "Data")
            {
                // Load all data for the required sheets.
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            }
            else
            {
                // Load only the structure for other sheets (no cell data).
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook.
            string sourcePath = "input.xlsx";

            // Configure load options with the custom filter.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new CustomLoadFilter();

            // Load the workbook using the specified load options.
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Display loaded worksheets and their cell counts.
            Console.WriteLine("Loaded worksheets:");
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"- {ws.Name}: Cells count = {ws.Cells.Count}");
            }

            // Save the workbook (optional, to verify the result).
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
