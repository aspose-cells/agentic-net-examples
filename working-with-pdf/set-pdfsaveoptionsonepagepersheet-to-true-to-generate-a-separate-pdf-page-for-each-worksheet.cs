using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Populate first worksheet with sample data
        Worksheet sheet1 = workbook.Worksheets[0];
        for (int i = 0; i < 10; i++)
        {
            sheet1.Cells[i, 0].Value = $"Sheet1 Data {i + 1}";
        }

        // Add a second worksheet and populate it
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        for (int i = 0; i < 10; i++)
        {
            sheet2.Cells[i, 0].Value = $"Sheet2 Data {i + 1}";
        }

        // Set PDF save options to generate a separate page for each worksheet
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true
        };

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}

// Author: Aspose.Cells .NET example code.