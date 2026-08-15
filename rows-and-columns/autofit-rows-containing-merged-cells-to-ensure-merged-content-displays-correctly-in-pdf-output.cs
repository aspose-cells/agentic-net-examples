// Title: C# – Auto‑fit rows with merged cells for accurate PDF output using Aspose.Cells
// Description: Shows how to merge a range, enable text wrapping, configure AutoFitterOptions (EachLine, Paragraph, ForRendering) and auto‑fit rows so merged content is rendered correctly when the workbook is saved as a PDF.
// Keywords: Aspose.Cells | C# | AutoFitRows | merged cells | PDF export | AutoFitterOptions | wrap text | row height | Excel to PDF | auto fit merged cells
// Common Searches: Aspose.Cells auto fit rows merged cells PDF | C# auto fit merged cells before PDF export | adjust row height for merged cells Aspose.Cells | AutoFitterOptions EachLine Paragraph ForRendering example | save merged cells with wrapped text to PDF using Aspose.Cells
// Developer Intent: Automatically adjust row heights for merged cells that contain wrapped text so the PDF rendering matches the worksheet layout.
// Use Cases: Generating a PDF report where a title spans several columns and must retain proper row height. | Creating invoices with merged header cells that hold long descriptions, ensuring the PDF shows the full text. | Exporting dashboards where merged cells contain paragraph‑style notes, requiring auto‑fit rows for readability.
// AI Prompts: Provide a C# example that auto‑fits rows with merged cells and wrapped text before saving to PDF using Aspose.Cells. | Explain how AutoFitterOptions EachLine, Paragraph, and ForRendering affect row height for merged cells in Aspose.Cells. | Show the steps to merge cells, enable text wrapping, and auto‑fit rows so the PDF output displays the merged content correctly.

using System;
using Aspose.Cells;

// Shows how to merge a range, enable text wrapping, configure AutoFitterOptions (EachLine, Paragraph, ForRendering) and auto‑fit rows so merged content is rendered correctly when the workbook is saved as a PDF.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add long text to a cell and enable text wrapping
        worksheet.Cells["A1"].PutValue("This is a long text that should be displayed correctly after auto‑fitting merged cells in PDF output.");
        Style style = worksheet.Cells["A1"].GetStyle();
        style.IsTextWrapped = true;
        worksheet.Cells["A1"].SetStyle(style);

        // Merge a range of cells (A1:C3)
        worksheet.Cells.Merge(0, 0, 3, 3);

        // Configure auto‑fitter options to handle merged cells
        AutoFitterOptions options = new AutoFitterOptions
        {
            AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine,
            AutoFitWrappedTextType = AutoFitWrappedTextType.Paragraph,
            ForRendering = true
        };

        // Auto‑fit rows considering merged cells
        worksheet.AutoFitRows(options);

        // Save the workbook as PDF
        workbook.Save("MergedAutoFit.pdf", SaveFormat.Pdf);
    }
}
