// Title: Export a merged workbook with charts and images to PDF using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook containing a column chart, optionally add a picture in a second workbook, merge the two files, configure PdfSaveOptions (ExportDocumentStructure, CalculateFormula), and save the combined workbook as a PDF for visual‑fidelity verification.
// Keywords: Aspose.Cells PDF export | C# merge workbooks | export chart to PDF | Aspose.Cells PdfSaveOptions | combined workbook PDF | ExportDocumentStructure | .NET Excel to PDF | chart image preservation
// Common Searches: Aspose.Cells combine workbooks and save as PDF C# | export Excel chart to PDF with Aspose.Cells | PdfSaveOptions ExportDocumentStructure example | merge two Excel files and generate PDF using Aspose | preserve images when converting Excel to PDF .NET
// Developer Intent: Merge multiple Excel workbooks and generate a single PDF that retains charts, images, and calculated formulas.
// Use Cases: Create a PDF report from several Excel sources where charts and pictures must appear exactly as in the original files. | Automate regression testing by exporting merged workbooks to PDF and comparing the output with baseline documents. | Generate printable documentation that combines data, visualizations, and embedded graphics from separate spreadsheets.
// AI Prompts: Write C# code that merges three Excel workbooks, each containing different chart types, and exports the result to a PDF with document structure and formula calculation enabled. | Show how to use Aspose.Cells PdfSaveOptions to keep images positioned correctly when converting a merged workbook to PDF. | Explain a method for programmatically validating the visual fidelity of charts in the exported PDF by comparing rendered pages to reference images.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook containing a column chart, optionally add a picture in a second workbook, merge the two files, configure PdfSaveOptions (ExportDocumentStructure, CalculateFormula), and save the combined workbook as a PDF for visual‑fidelity verification.
class ExportCombinedWorkbookToPdf
{
    static void Main()
    {
        // ---------- Create first workbook and add a chart ----------
        Workbook wb1 = new Workbook();                                   // create workbook
        Worksheet ws1 = wb1.Worksheets[0];
        ws1.Name = "DataSheet";

        // Populate sample data
        ws1.Cells["A1"].PutValue("Category");
        ws1.Cells["A2"].PutValue("Fruits");
        ws1.Cells["A3"].PutValue("Vegetables");
        ws1.Cells["B1"].PutValue("Value");
        ws1.Cells["B2"].PutValue(50);
        ws1.Cells["B3"].PutValue(30);

        // Add a column chart linked to the data
        int chartIdx = ws1.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = ws1.Charts[chartIdx];
        chart.NSeries.Add("B2:B3", true);               // values
        chart.NSeries.CategoryData = "A2:A3";           // categories
        chart.Title.Text = "Sample Chart";

        // ---------- Create second workbook (could contain images, shapes, etc.) ----------
        Workbook wb2 = new Workbook();                                   // create second workbook
        Worksheet ws2 = wb2.Worksheets[0];
        ws2.Name = "ImageSheet";

        // Example: insert a picture (requires an existing image file)
        // ws2.Pictures.Add(1, 1, "sample_image.png"); // Uncomment and provide a valid path if needed

        // ---------- Combine the two workbooks ----------
        wb1.Combine(wb2);                                                // combine second into first

        // ---------- Configure PDF save options ----------
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.ExportDocumentStructure = true;   // retain document structure for verification
        pdfOptions.CalculateFormula = true;         // ensure any formulas are evaluated

        // ---------- Save the combined workbook as PDF ----------
        wb1.Save("CombinedWorkbook.pdf", pdfOptions); // save using the Save(string, SaveOptions) rule
    }
}
