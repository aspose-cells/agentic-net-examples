// Title: Load Only the First Worksheet from a Large Excel File with Aspose.Cells (C#)
// Description: Shows how to implement a custom LoadFilter that loads just the first worksheet (index 0) via LoadOptions, creates a Workbook with this filter, and saves the resulting single‑sheet file.
// Keywords: Aspose.Cells | C# | LoadFilter | first worksheet | partial workbook loading | large Excel file | memory optimization | exclude sheets
// Common Searches: Aspose.Cells load first sheet only | C# load specific worksheet without loading whole workbook | How to use LoadFilter to load selected sheets | Reduce memory when opening large Excel with Aspose | Save workbook after loading selected worksheets
// Developer Intent: Load only the first worksheet from a large workbook and save it as a new file while minimizing memory consumption.
// Use Cases: Process massive workbooks by loading just the first sheet to lower RAM usage. | Create a lightweight copy that contains only the initial worksheet for downstream analysis. | Quickly preview or export the first sheet of a multi‑sheet file without the overhead of loading the rest.
// AI Prompts: Generate C# code that loads the first three worksheets using a custom LoadFilter. | Show how to load a worksheet by its name with Aspose.Cells LoadFilter in C#. | Explain how to combine LoadFilter and LoadOptions to load selected sheets and export them to PDF.

using System;
using Aspose.Cells;

namespace LoadFirstWorksheetExample
{
    // Custom LoadFilter that loads only the first worksheet (index 0)
    // Shows how to implement a custom LoadFilter that loads just the first worksheet (index 0) via LoadOptions, creates a Workbook with this filter, and saves the resulting single‑sheet file.
    public class FirstSheetOnlyLoadFilter : LoadFilter
    {
        // Define the order of sheets to load – only the first sheet
        private readonly int[] _sheetsOrder = new int[] { 0 };

        // Override the read‑only property to return our custom order
        public override int[] SheetsInLoadingOrder => _sheetsOrder;
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook (large file with many sheets)
            string sourcePath = "LargeWorkbook.xlsx";

            // Configure LoadOptions with the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new FirstSheetOnlyLoadFilter();

            // Load the workbook using the constructor that accepts a file name and LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // At this point only the first worksheet is loaded; other sheets are absent
            Console.WriteLine("Loaded worksheets count: " + workbook.Worksheets.Count);
            Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);

            // Save the trimmed workbook to a new file
            string outputPath = "FirstSheetOnly.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine("Workbook saved with only the first sheet to: " + outputPath);
        }
    }
}
