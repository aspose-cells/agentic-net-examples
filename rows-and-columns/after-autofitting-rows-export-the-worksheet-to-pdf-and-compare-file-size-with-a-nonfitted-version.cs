// Title: Export a worksheet with AutoFitRows to PDF and compare its file size to a non‑auto‑fitted PDF using Aspose.Cells for .NET
// Common Searches: how to use Aspose.Cells to auto‑fit rows before saving as PDF in C# | C# compare PDF file size of auto‑fitted rows versus original worksheet with Aspose.Cells | Aspose.Cells AutoFitRows impact on PDF output size | measure difference in PDF size after applying AutoFitRows in Aspose.Cells | export Excel to PDF with row auto‑fit and get file size in C#
// Tags: AutoFitRows PDF export Aspose.Cells | PDF size comparison after row auto‑fit Aspose.Cells | C# Aspose.Cells row auto‑fit impact on PDF size | export worksheet to PDF with AutoFitRows Aspose.Cells | measure PDF file size difference Aspose.Cells row fitting

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, auto‑fits all rows in the first worksheet, saves it as a PDF, then saves a second PDF from the original workbook without auto‑fitting, and finally prints the byte sizes of both PDFs and their difference.
class Program
{
    static void Main()
    {
        // Paths for the source Excel file and the resulting PDFs
        string excelPath = "input.xlsx";
        string pdfFittedPath = "output_fitted.pdf";
        string pdfNonFittedPath = "output_nonfitted.pdf";

        // Load the workbook and auto‑fit all rows in the first worksheet
        Workbook wbFitted = new Workbook(excelPath);
        Worksheet ws = wbFitted.Worksheets[0];
        ws.AutoFitRows(); // Auto‑fit rows

        // Export the auto‑fitted worksheet to PDF
        wbFitted.Save(pdfFittedPath, SaveFormat.Pdf);

        // Load the original workbook again (without auto‑fit) and export to PDF
        Workbook wbNonFitted = new Workbook(excelPath);
        wbNonFitted.Save(pdfNonFittedPath, SaveFormat.Pdf);

        // Compare file sizes
        long sizeFitted = new FileInfo(pdfFittedPath).Length;
        long sizeNonFitted = new FileInfo(pdfNonFittedPath).Length;

        Console.WriteLine($"Fitted PDF size: {sizeFitted} bytes");
        Console.WriteLine($"Non‑fitted PDF size: {sizeNonFitted} bytes");
        Console.WriteLine($"Size difference: {sizeFitted - sizeNonFitted} bytes");
    }
}
