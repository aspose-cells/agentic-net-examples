using System;
using Aspose.Cells;

class ExportWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook (or load an existing one with XML maps)
        Workbook workbook = new Workbook();

        // Populate some sample data to illustrate mapped cells
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("John");
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Alice");

        // Configure PDF save options to preserve visual representation
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.ExportDocumentStructure = true;   // retain document structure (mapped cells)
        pdfOptions.CalculateFormula = true;         // ensure formulas are evaluated before export

        // Save the workbook to PDF using the provided Save(string, SaveOptions) rule
        workbook.Save("MappedWorkbook.pdf", pdfOptions);
    }
}