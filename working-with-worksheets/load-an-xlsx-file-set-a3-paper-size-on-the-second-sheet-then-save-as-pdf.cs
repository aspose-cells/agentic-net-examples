using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourceFile = "input.xlsx";

        // Path for the resulting PDF file
        string pdfFile = "output.pdf";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourceFile);

        // Access the second worksheet (index 1) and set its paper size to A3
        Worksheet secondSheet = workbook.Worksheets[1];
        secondSheet.PageSetup.PaperSize = PaperSizeType.PaperA3;

        // Save the modified workbook as a PDF document
        workbook.Save(pdfFile, SaveFormat.Pdf);
    }
}