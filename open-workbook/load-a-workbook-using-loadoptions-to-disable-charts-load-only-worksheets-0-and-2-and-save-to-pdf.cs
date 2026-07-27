// Title: Load selected worksheets (0 & 2) without charts using Aspose.Cells LoadOptions and export to PDF (C#)
// Description: Demonstrates how to create a custom LoadFilter that disables chart loading and limits the workbook to sheet indexes 0 and 2. The filter is applied through LoadOptions when constructing a Workbook, and the resulting workbook is saved as a PDF, yielding a lightweight document that contains only the chosen sheets and no chart objects.
// Keywords: Aspose.Cells | LoadOptions | LoadFilter | C# | disable charts | select worksheets | sheet indexes | export to PDF | Excel to PDF | filter workbook
// Common Searches: Aspose.Cells load only specific sheets | How to exclude charts when loading an Excel file with Aspose | C# load workbook with LoadOptions and save as PDF | LoadFilter example for sheet selection in Aspose.Cells | Export selected worksheets to PDF using Aspose.Cells
// Developer Intent: Load an Excel file while skipping all charts and loading only sheets 0 and 2, then convert the filtered workbook to PDF.
// Use Cases: Create a compact PDF report that includes only summary and data sheets, omitting large chart objects. | Reduce processing time and memory usage by loading just the required worksheets from a massive workbook. | Provide users with a PDF preview of selected Excel tabs without rendering embedded charts.
// AI Prompts: Generate C# code that uses a custom LoadFilter in Aspose.Cells to load only sheet indexes 0 and 2 and exclude charts, then save the workbook as PDF. | Explain how LoadOptions and LoadFilter work together to filter worksheets and chart objects in Aspose.Cells. | Show step‑by‑step instructions for loading an Excel file with selected sheets and no charts, and exporting it to PDF using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadFilterExample
{
    // Custom LoadFilter to control which data is loaded and which sheets are loaded
    // Demonstrates how to create a custom LoadFilter that disables chart loading and limits the workbook to sheet indexes 0 and 2. The filter is applied through LoadOptions when constructing a Workbook, and the resulting workbook is saved as a PDF, yielding a lightweight document that contains only the chosen sheets and no chart objects.
    public class CustomLoadFilter : LoadFilter
    {
        // Load all data except charts
        private const LoadDataFilterOptions LoadOptionsWithoutCharts =
            LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart;

        // Specify the sheet indexes to load (0‑based)
        private static readonly int[] SheetsToLoad = new int[] { 0, 2 };

        public CustomLoadFilter() : base(LoadOptionsWithoutCharts)
        {
        }

        // Override to provide the custom sheet loading order
        public override int[] SheetsInLoadingOrder => SheetsToLoad;
    }

    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Configure LoadOptions with the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new CustomLoadFilter();

            // Load the workbook using the constructor that accepts a file name and LoadOptions
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the loaded workbook as PDF (charts are omitted, only sheets 0 and 2 are present)
            string outputPath = "output.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook loaded with sheets 0 and 2 (charts excluded) and saved to PDF at '{outputPath}'.");
        }
    }
}
