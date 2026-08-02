// Title: Skip Hidden Worksheets with a Custom LoadFilter in Aspose.Cells for .NET
// Description: Shows how to implement a VisibleSheetsLoadFilter that overrides StartSheet to load only worksheets where IsVisible is true, assign the filter to LoadOptions, open a workbook, enumerate the loaded sheets, and optionally save the filtered file.
// Keywords: Aspose.Cells | .NET | LoadFilter | visible worksheets | hidden sheets | LoadOptions | Workbook loading | C# example | skip hidden sheets | custom filter
// Common Searches: Aspose.Cells load only visible sheets | ignore hidden worksheets when opening a workbook C# | custom LoadFilter example Aspose.Cells | skip hidden worksheets Aspose.Cells .NET | LoadOptions LoadFilter usage
// Developer Intent: Open a workbook while automatically excluding hidden worksheets by applying a custom LoadFilter.
// Use Cases: Improve performance when processing large workbooks by loading only visible sheets. | Create a copy of a workbook that contains just the visible worksheets. | Generate reports that require only user‑visible data without hidden tabs. | Validate workbook content by enumerating loaded sheets and their visibility.
// AI Prompts: Generate a C# class that inherits from Aspose.Cells.LoadFilter and skips hidden worksheets. | Show how to set LoadOptions.LoadFilter to the custom filter and open a workbook. | Provide code to list the names and visibility of worksheets after loading with the filter. | Explain why overriding StartSheet is sufficient to exclude hidden sheets.

using System;
using Aspose.Cells;

namespace Example
{
    // Custom LoadFilter that loads only visible worksheets
    // Shows how to implement a VisibleSheetsLoadFilter that overrides StartSheet to load only worksheets where IsVisible is true, assign the filter to LoadOptions, open a workbook, enumerate the loaded sheets, and optionally save the filtered file.
    public class VisibleSheetsLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load data only if the sheet is visible
            if (sheet.IsVisible)
            {
                // Proceed with default loading behavior
                base.StartSheet(sheet);
            }
            // If the sheet is hidden, do not call base.StartSheet, effectively skipping it
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the source workbook
            string sourcePath = "Input.xlsx";

            // Create LoadOptions and assign the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new VisibleSheetsLoadFilter();

            // Load the workbook with the filter applied
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Display loaded worksheets and their visibility status
            Console.WriteLine("Loaded worksheets:");
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"- {ws.Name} (Visible: {ws.IsVisible})");
            }

            // Save the workbook (optional)
            workbook.Save("FilteredOutput.xlsx");
        }
    }
}
