// Title: Compare PDF output file sizes when exporting an Excel workbook with Aspose.Cells using AllColumnsInOnePagePerSheet versus default settings in C#
// AI Prompts: Generate a C# console program that loads an .xlsx file, saves it to PDF twice—once with default SaveFormat.Pdf and once with PdfSaveOptions.AllColumnsInOnePagePerSheet set to true—and prints the byte size of each PDF. | Modify existing Aspose.Cells PDF export code to record the size difference between the default PDF and the AllColumnsInOnePagePerSheet PDF, then output a summary indicating which file is larger. | Create a C# script that iterates over multiple Excel workbooks, exports each to PDF using both default settings and the AllColumnsInOnePagePerSheet option, and logs a table of file size comparisons.
// Common Searches: how does PdfSaveOptions.AllColumnsInOnePagePerSheet affect PDF file size in Aspose.Cells C# | C# Aspose.Cells compare default PDF export size with AllColumnsInOnePagePerSheet option | measure generated PDF size after exporting Excel workbook with Aspose.Cells | Aspose.Cells PDF export size differences between default and AllColumnsInOnePagePerSheet | C# code to get file size of PDFs created by Aspose.Cells
// Tags: Aspose.Cells PDF export AllColumnsInOnePagePerSheet | C# PDF size comparison Aspose.Cells | PdfSaveOptions file size impact | Excel to PDF size analysis Aspose.Cells | PDF size measurement C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Loads an Excel workbook, saves it to PDF using default settings and using PdfSaveOptions with AllColumnsInOnePagePerSheet=true, then outputs and compares the resulting file sizes.
class PdfSizeComparison
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        string excelPath = "input.xlsx";
        Workbook workbook = new Workbook(excelPath);

        // Define output PDF file names
        string defaultPdfPath = "default.pdf";
        string allColumnsPdfPath = "AllColumnsInOnePagePerSheet.pdf";

        // ---------- Default PDF export ----------
        // Save the workbook to PDF using default settings
        workbook.Save(defaultPdfPath, SaveFormat.Pdf);

        // Get file size of the default PDF
        long defaultSize = new FileInfo(defaultPdfPath).Length;

        // ---------- PDF export with AllColumnsInOnePagePerSheet ----------
        // Configure PDF save options to fit all columns on one page per sheet
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            AllColumnsInOnePagePerSheet = true
        };

        // Save the workbook to PDF with the specified option
        workbook.Save(allColumnsPdfPath, pdfOptions);

        // Get file size of the PDF with AllColumnsInOnePagePerSheet option
        long allColumnsSize = new FileInfo(allColumnsPdfPath).Length;

        // ---------- Comparison ----------
        Console.WriteLine($"Default PDF size: {defaultSize} bytes");
        Console.WriteLine($"AllColumnsInOnePagePerSheet PDF size: {allColumnsSize} bytes");

        if (defaultSize == allColumnsSize)
        {
            Console.WriteLine("Both PDFs have the same file size.");
        }
        else if (defaultSize > allColumnsSize)
        {
            Console.WriteLine("The default PDF is larger.");
        }
        else
        {
            Console.WriteLine("The AllColumnsInOnePagePerSheet PDF is larger.");
        }
    }
}
