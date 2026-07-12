using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (empty by default)
        Workbook workbook = new Workbook();

        // Example: add some data to the first worksheet (optional)
        Worksheet sheet = workbook.Worksheets[0];
        for (int i = 0; i < 10; i++)
        {
            sheet.Cells[i, 0].Value = $"Row {i + 1}";
        }

        // Configure PDF save options to force each worksheet onto a single page
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true // each sheet -> one PDF page
        };

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}

// Author: Aspose.Cells .NET example code.