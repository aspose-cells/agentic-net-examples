using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        string sourceFile = "input.xlsx";
        Workbook workbook = new Workbook(sourceFile);

        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        string pdfFile = "output.pdf";
        workbook.Save(pdfFile, pdfOptions);
    }
}