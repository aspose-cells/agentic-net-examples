// Title: Remove all charts from an Excel workbook and export it to PDF while preserving cell formatting with Aspose.Cells for .NET
// Common Searches: how to delete every chart in an Excel file using Aspose.Cells C# | export Excel to PDF without charts preserving cell styles Aspose.Cells | Aspose.Cells C# remove worksheet charts before PDF conversion | save chart‑free workbook as PDF keeping original formatting in .NET | remove charts from workbook then convert to PDF with Aspose.Cells
// Tags: Aspose.Cells remove charts from workbook | PDF conversion preserving cell formatting Aspose.Cells | chart‑free Excel to PDF export | Aspose.Cells PdfSaveOptions cell style retention | C# delete worksheet charts Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load the existing workbook (replace with your actual file path)
Workbook workbook = new Workbook("input.xlsx");

// Remove all charts from each worksheet to make the workbook chart‑free
foreach (Worksheet sheet in workbook.Worksheets)
{
    // Collect chart indexes because removing charts shifts the collection
    int chartCount = sheet.Charts.Count;
    for (int i = chartCount - 1; i >= 0; i--)
    {
        sheet.Charts.RemoveAt(i);
    }
}

// Configure PDF save options to preserve original cell formatting
PdfSaveOptions pdfOptions = new PdfSaveOptions
{
    // Keep the default behavior which retains cell styles, fonts, colors, etc.
    // Additional options can be set here if needed, e.g., one page per sheet:
    // OnePagePerSheet = false
};

// Save the chart‑free workbook as a PDF document
workbook.Save("output.pdf", pdfOptions);
