// Title: Reload an Excel workbook, change printer paper size to A3, and compare PDF output sizes for A5 vs A3 using Aspose.Cells in C#
// Common Searches: how to change worksheet printer paper size to A3 with Aspose.Cells before saving as PDF in C# | compare PDF file size when exporting Excel to A5 and A3 paper sizes using Aspose.Cells | reload an existing workbook and switch page setup for multiple PDF exports in C# Aspose.Cells | measure difference in PDF file size between A5 and A3 page setups with Aspose.Cells | C# code to get file size of PDFs generated from different paper sizes using Aspose.Cells
// Tags: worksheet page setup paper size A3 Aspose.Cells | export workbook to PDF with custom paper size Aspose.Cells | compare PDF output sizes from A5 and A3 paper formats C# | reload workbook and modify page setup Aspose.Cells | measure PDF file size after changing printer paper size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, saves it as a PDF using A5 printer paper size, records the file size, then changes the printer paper size to A3, saves another PDF, records that size, and prints a comparison indicating which PDF is larger.
class ComparePdfSizes
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Define output PDF file names
        string pdfA5Path = "output_A5.pdf";
        string pdfA3Path = "output_A3.pdf";

        // Set printer paper size to A5 for all worksheets and save as PDF
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA5;
        }
        workbook.Save(pdfA5Path, SaveFormat.Pdf);

        // Get file size of A5 PDF
        long sizeA5 = new FileInfo(pdfA5Path).Length;

        // Change printer paper size to A3 for all worksheets and save as PDF
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
        }
        workbook.Save(pdfA3Path, SaveFormat.Pdf);

        // Get file size of A3 PDF
        long sizeA3 = new FileInfo(pdfA3Path).Length;

        // Compare the two PDF file sizes and output the result
        Console.WriteLine($"A5 PDF size: {sizeA5} bytes");
        Console.WriteLine($"A3 PDF size: {sizeA3} bytes");

        if (sizeA3 > sizeA5)
        {
            Console.WriteLine("A3 PDF is larger than A5 PDF.");
        }
        else if (sizeA3 < sizeA5)
        {
            Console.WriteLine("A3 PDF is smaller than A5 PDF.");
        }
        else
        {
            Console.WriteLine("A3 PDF and A5 PDF have the same size.");
        }
    }
}
