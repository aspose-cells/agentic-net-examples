// Title: Compare PDF file sizes of the same workbook using A5 and A3 printer paper settings with Aspose.Cells for .NET
// AI Prompts: Write a C# console application that loads an Excel workbook, sets the first worksheet's PageSetup.PaperSize to PaperA5, saves it as a PDF, then changes the PaperSize to PaperA3, saves a second PDF, and prints the byte size of each file. | Update the program to reload or reset the workbook after changing PageSetup so the second PDF reflects the new paper size correctly. | Enhance the code to accept input Excel path and output PDF paths via command‑line arguments and log the size comparison result to a text file.
// Common Searches: how to change worksheet paper size to A3 when exporting Excel to PDF using Aspose.Cells C# | Aspose.Cells PDF output size difference between A5 and A3 page setups | C# code to compare byte size of PDFs generated from the same workbook with different paper sizes | reload workbook after modifying PageSetup in Aspose.Cells before saving another PDF
// Tags: Aspose.Cells set worksheet paper size A5 PDF | Aspose.Cells set worksheet paper size A3 PDF | Aspose.Cells compare PDF output size by paper size | C# get PDF file length Aspose.Cells | Aspose.Cells reload workbook after PageSetup change

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, exports the first worksheet to PDF with A5 paper size, then changes the paper size to A3, exports a second PDF, reads both file sizes, and prints which PDF is larger.
class PdfPaperSizeComparison
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Set printer paper size to A5 and save as PDF
        workbook.Worksheets[0].PageSetup.PaperSize = PaperSizeType.PaperA5;
        workbook.Save("output_A5.pdf", SaveFormat.Pdf);

        // Set printer paper size to A3 and save as PDF
        workbook.Worksheets[0].PageSetup.PaperSize = PaperSizeType.PaperA3;
        workbook.Save("output_A3.pdf", SaveFormat.Pdf);

        // Get file sizes
        long sizeA5 = new FileInfo("output_A5.pdf").Length;
        long sizeA3 = new FileInfo("output_A3.pdf").Length;

        // Output comparison result
        Console.WriteLine($"A5 PDF size: {sizeA5} bytes");
        Console.WriteLine($"A3 PDF size: {sizeA3} bytes");
        Console.WriteLine(sizeA3 > sizeA5
            ? "A3 PDF is larger than A5 PDF."
            : sizeA3 < sizeA5
                ? "A3 PDF is smaller than A5 PDF."
                : "A3 PDF and A5 PDF have the same size.");
    }
}
