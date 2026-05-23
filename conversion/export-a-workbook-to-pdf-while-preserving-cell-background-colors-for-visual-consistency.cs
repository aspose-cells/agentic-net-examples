using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ExportWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate cells with sample data and apply background colors
        worksheet.Cells["A1"].PutValue("Red Background");
        Style redStyle = workbook.CreateStyle();
        redStyle.ForegroundColor = Color.Red;
        redStyle.Pattern = BackgroundType.Solid;
        worksheet.Cells["A1"].SetStyle(redStyle);

        worksheet.Cells["B1"].PutValue("Green Background");
        Style greenStyle = workbook.CreateStyle();
        greenStyle.ForegroundColor = Color.Green;
        greenStyle.Pattern = BackgroundType.Solid;
        worksheet.Cells["B1"].SetStyle(greenStyle);

        // Create PDF save options to retain document structure and visual fidelity
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.ExportDocumentStructure = true; // preserve structure for consistency

        // Ensure any formulas are calculated before saving
        workbook.CalculateFormula();

        // Save the workbook as PDF using the provided save method with options
        workbook.Save("ColoredWorkbook.pdf", pdfOptions);
    }
}