using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // for PdfSaveOptions

// Author: Aspose.Cells .NET example – demonstrates that PDF respects TextCrossType (CrossStringInPdf) setting
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put a long string that will overflow the cell width
        sheet.Cells["A1"].PutValue("This is a very long text that should cross cell boundaries when rendered to PDF.");

        // Reduce column width so the text definitely exceeds the cell width
        sheet.Cells.SetColumnWidth(0, 5); // column A width = 5 characters

        // Configure PDF save options to use the CrossKeep behavior (cross‑string in PDF)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            TextCrossType = TextCrossType.CrossKeep // ensures the string crosses into adjacent cells
        };

        // Save the workbook as PDF
        workbook.Save("CrossStringDemo.pdf", pdfOptions);

        Console.WriteLine("PDF saved with TextCrossType.CrossKeep – verify cross‑string behavior in the output file.");
    }
}