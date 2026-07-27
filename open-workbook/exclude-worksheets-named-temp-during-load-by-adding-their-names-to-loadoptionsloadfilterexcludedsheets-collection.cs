// Title: Exclude "Temp" worksheets when loading a workbook with Aspose.Cells LoadOptions (C#)
// Description: Demonstrates how to prevent worksheets named "Temp" from being loaded by configuring LoadOptions in Aspose.Cells for .NET. A custom LoadFilter sets LoadDataFilterOptions.None for matching sheets, and the example shows loading, enumerating, and optionally saving the filtered workbook.
// Keywords: Aspose.Cells C# | .NET LoadOptions | exclude worksheet on load | skip Temp sheet | custom LoadFilter | memory optimization Aspose.Cells | LoadOptions.ExcludedSheets | Excel automation | US developers | global
// Common Searches: How to skip specific sheets when opening an Excel file with Aspose.Cells .NET | Aspose.Cells LoadOptions exclude Temp worksheet | C# load workbook without temporary sheets | Custom LoadFilter example Aspose.Cells | LoadOptions.ExcludedSheets usage
// Developer Intent: Load an Excel workbook while automatically omitting any worksheet whose name is "Temp" using Aspose.Cells LoadOptions in C#.
// Use Cases: Reduce memory usage by ignoring large temporary worksheets during bulk processing. | Process only business‑critical sheets in a multi‑sheet workbook without extra filtering code. | Create a clean copy of a workbook that excludes temporary or staging sheets before distribution.
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions with a custom LoadFilter to exclude worksheets named "Temp" and then saves the resulting workbook. | Explain why setting LoadDataFilterOptions.None in a LoadFilter prevents a sheet from being loaded and show how to extend the filter for multiple sheet names. | Provide a concise solution that adds "Temp" to LoadOptions.LoadFilter.ExcludedSheets collection instead of a custom filter.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Custom load filter that skips worksheets named "Temp"
    // Demonstrates how to prevent worksheets named "Temp" from being loaded by configuring LoadOptions in Aspose.Cells for .NET. A custom LoadFilter sets LoadDataFilterOptions.None for matching sheets, and the example shows loading, enumerating, and optionally saving the filtered workbook.
    public class ExcludeTempSheetsFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // If the worksheet name is "Temp", do not load any data for it
            if (sheet.Name.Equals("Temp", StringComparison.OrdinalIgnoreCase))
            {
                // Skip loading all data for this sheet
                LoadDataFilterOptions = LoadDataFilterOptions.None;
            }
            else
            {
                // Load everything for other sheets
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "InputWorkbook.xlsx";

            // Create LoadOptions and assign the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new ExcludeTempSheetsFilter();

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // At this point, worksheets named "Temp" are excluded from loading
            // You can verify by iterating through the loaded worksheets
            Console.WriteLine("Loaded worksheets:");
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"- {ws.Name}");
            }

            // Save the workbook to a new file (optional)
            string outputPath = "OutputWorkbook.xlsx";
            workbook.Save(outputPath);
        }
    }
}
