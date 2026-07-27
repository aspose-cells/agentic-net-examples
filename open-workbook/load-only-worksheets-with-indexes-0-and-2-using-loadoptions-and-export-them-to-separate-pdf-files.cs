// Title: Load Selected Worksheets (0 & 2) with LoadOptions and Export Each to PDF – Aspose.Cells for .NET
// Description: Demonstrates how to create a custom LoadFilter that loads only worksheet indexes 0 and 2, apply it via LoadOptions, and then save each loaded sheet as an individual PDF using PdfSaveOptions and SheetSet. This approach reduces memory usage and speeds up conversion when only specific sheets are needed.
// Keywords: Aspose.Cells | .NET | LoadOptions | LoadFilter | select worksheets | export to PDF | sheet index | PdfSaveOptions | custom filter | partial workbook loading | C# Excel to PDF
// Common Searches: Aspose.Cells load specific worksheets | Load only certain sheets from Excel .NET | Export individual worksheets to PDF using Aspose.Cells | How to use LoadFilter with LoadOptions in C# | Save selected Excel sheets as separate PDF files
// Developer Intent: Load only worksheets 0 and 2 from an Excel workbook and generate a separate PDF file for each sheet using Aspose.Cells for .NET.
// Use Cases: Create per‑sheet PDF reports when only a subset of worksheets contains publishable data. | Reduce memory consumption in batch conversions by loading only the required sheets. | Automate archival of selected worksheets as PDFs without processing the entire workbook.
// AI Prompts: Show how to change the code to load worksheets 1 and 3 instead of 0 and 2. | Provide an example of setting PdfSaveOptions to use landscape orientation for each exported PDF. | Explain how to retrieve the number of worksheets actually loaded after applying a custom LoadFilter.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Saving;

// Custom LoadFilter that loads only the specified sheet indexes in the given order
// Demonstrates how to create a custom LoadFilter that loads only worksheet indexes 0 and 2, apply it via LoadOptions, and then save each loaded sheet as an individual PDF using PdfSaveOptions and SheetSet. This approach reduces memory usage and speeds up conversion when only specific sheets are needed.
public class CustomLoadFilter : LoadFilter
{
    private readonly int[] _sheetsOrder;

    public CustomLoadFilter(int[] sheetsOrder) : base(LoadDataFilterOptions.All)
    {
        _sheetsOrder = sheetsOrder;
    }

    // Override to return the desired sheet indexes
    public override int[] SheetsInLoadingOrder => _sheetsOrder;
}

class Program
{
    static void Main()
    {
        // Path to the source workbook
        string sourceFile = "input.xlsx";

        // Create LoadOptions and assign the custom filter to load only sheets 0 and 2
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new CustomLoadFilter(new int[] { 0, 2 });

        // Load the workbook with the specified options (only the selected sheets are loaded)
        Workbook workbook = new Workbook(sourceFile, loadOptions);

        // Export each loaded worksheet to a separate PDF file
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            // Prepare PDF save options to render only the current worksheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.SheetSet = new SheetSet(new int[] { i }); // render sheet at index i

            // Build output file name using the worksheet name for clarity
            string sheetName = workbook.Worksheets[i].Name;
            string outputPdf = $"Sheet_{sheetName}.pdf";

            // Save the workbook (only the selected sheet) as PDF
            workbook.Save(outputPdf, pdfOptions);
        }
    }
}
