using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["B1"].PutValue("World");

        // Add a comment to cell A1
        int commentIndex = worksheet.Comments.Add("A1");
        worksheet.Comments[commentIndex].Note = "This is a comment that will appear as a PDF annotation.";

        // Create PDF save options and enable document structure export
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.ExportDocumentStructure = true; // Preserves cell comments as PDF annotations

        // Calculate formulas (if any) before saving
        workbook.CalculateFormula();

        // Save the workbook as PDF with the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}