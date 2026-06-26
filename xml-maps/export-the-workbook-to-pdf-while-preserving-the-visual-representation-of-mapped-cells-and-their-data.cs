using System;
using Aspose.Cells;

class ExportWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (visual representation of mapped cells)
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Alice");
        sheet.Cells["B3"].PutValue(25);

        // Ensure any formulas are calculated before saving
        workbook.CalculateFormula();

        // Configure PDF save options to preserve document structure
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.ExportDocumentStructure = true; // retains visual mapping of cells

        // Save the workbook to PDF using the provided save rule
        workbook.Save("MappedWorkbook.pdf", pdfOptions);
    }
}