using System;
using System.IO;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – calculate formulas before PDF export
class Program
{
    static void Main()
    {
        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate cells with sample data and a formula
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        worksheet.Cells["A3"].Formula = "=SUM(A1:A2)";

        // Explicitly calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Configure PDF save options to also calculate formulas (ensures safety)
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            CalculateFormula = true
        };

        // Save the workbook as a PDF file
        string outputFile = "Result.pdf";
        workbook.Save(outputFile, pdfSaveOptions);
    }
}