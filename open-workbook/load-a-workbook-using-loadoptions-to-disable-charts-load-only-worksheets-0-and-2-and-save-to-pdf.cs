// Title: Load selected worksheets without charts using Aspose.Cells LoadOptions and export to PDF (C#)
// Description: Demonstrates how to create a LoadOptions object with a custom LoadFilter that disables chart loading and loads only worksheet indexes 0 and 2, then opens an Excel file and saves the resulting workbook as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | LoadOptions | CustomLoadFilter | disable chart loading | load specific worksheets | C# PDF export | Aspose.Cells .NET | select sheets by index | Excel to PDF without charts
// Common Searches: How can I load only certain sheets with Aspose.Cells and skip charts? | Aspose.Cells C# load worksheets 0 and 2 only | Export selected Excel sheets to PDF using LoadOptions | Disable chart loading in Aspose.Cells to improve performance | Custom LoadFilter example for Aspose.Cells .NET
// Developer Intent: Load an Excel workbook while excluding chart objects and loading only the first and third worksheets, then convert the workbook to PDF.
// Use Cases: Generate a lightweight PDF report that contains only data tables, omitting chart graphics to reduce file size. | Speed up processing of large workbooks by loading only the required sheets and ignoring unnecessary chart data. | Archive specific worksheets of an Excel file as PDF while discarding visual chart elements.
// AI Prompts: Show how to modify the CustomLoadFilter to also exclude images while loading selected worksheets. | Provide code that loads worksheets 1 and 3 and saves each to separate PDF files using Aspose.Cells. | Explain how to configure LoadDataFilterOptions to load formulas but skip charts for particular sheets.

using System;
using Aspose.Cells;

// Demonstrates how to create a LoadOptions object with a custom LoadFilter that disables chart loading and loads only worksheet indexes 0 and 2, then opens an Excel file and saves the resulting workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create LoadOptions and assign a custom LoadFilter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new CustomLoadFilter();

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Save the loaded workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }

    // Custom LoadFilter to exclude charts and load only sheets 0 and 2
    class CustomLoadFilter : LoadFilter
    {
        // Initialize with all data options; we'll adjust per sheet
        public CustomLoadFilter() : base(LoadDataFilterOptions.All) { }

        // Called for each sheet being loaded
        public override void StartSheet(Worksheet sheet)
        {
            // Load everything except charts
            LoadDataFilterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart;
        }

        // Specify the exact sheet indexes to load (0‑based)
        public override int[] SheetsInLoadingOrder => new int[] { 0, 2 };
    }
}
