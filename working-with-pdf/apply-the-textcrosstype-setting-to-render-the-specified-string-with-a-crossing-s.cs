using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTextCrossTypePdfDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a long text into cell A1 that will exceed the column width
            worksheet.Cells["A1"].PutValue("This is a very long text that will cross cell boundaries when rendered to PDF.");

            // Reduce the column width to force the text to overflow
            worksheet.Cells.SetColumnWidth(0, 5); // Column A width set to 5 characters

            // Create PDF save options and set the desired TextCrossType
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.TextCrossType = TextCrossType.CrossKeep; // Text will cross other cells and keep their text

            // Save the workbook as a PDF using the configured options
            workbook.Save("TextCrossTypeDemo.pdf", pdfOptions);

            Console.WriteLine("PDF saved with TextCrossType.CrossKeep setting.");
        }
    }
}