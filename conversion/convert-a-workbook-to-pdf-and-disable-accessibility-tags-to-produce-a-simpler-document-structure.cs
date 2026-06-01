using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Alice");
        sheet.Cells["B3"].PutValue(25);

        // Create PDF save options and disable the export of document structure
        // (this removes accessibility tags for a simpler PDF)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.ExportDocumentStructure = false;

        // Ensure any formulas are calculated before saving
        workbook.CalculateFormula();

        // Save the workbook as a PDF file using the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}