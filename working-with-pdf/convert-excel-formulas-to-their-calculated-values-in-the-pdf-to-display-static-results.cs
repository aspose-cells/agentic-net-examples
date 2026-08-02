using System;
using System.IO;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – converts formulas to static values in PDF
class Program
{
    static void Main()
    {
        // Load the source Excel workbook (replace with your actual file path)
        string excelFilePath = "input.xlsx";
        Workbook workbook = new Workbook(excelFilePath);

        // Configure PDF save options to calculate formulas before rendering
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            // When true, all formulas are evaluated and their results are written to the PDF
            CalculateFormula = true
        };

        // Save the workbook as a PDF file with calculated formula results
        string pdfFilePath = "output.pdf";
        workbook.Save(pdfFilePath, pdfSaveOptions);
    }
}