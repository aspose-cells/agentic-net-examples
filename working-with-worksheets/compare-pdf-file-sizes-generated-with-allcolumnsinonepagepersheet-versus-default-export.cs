// Title: Compare PDF file sizes when using AllColumnsInOnePagePerSheet versus default export in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to measure PDF size with AllColumnsInOnePagePerSheet option | difference in PDF file size between default export and AllColumnsInOnePagePerSheet in Aspose.Cells | C# code to compare sizes of PDFs generated from Excel using PdfSaveOptions | Aspose.Cells PDF output size comparison for column fitting setting | evaluate PDF size impact of AllColumnsInOnePagePerSheet in .NET
// Tags: AllColumnsInOnePagePerSheet size impact | PdfSaveOptions column-to-page setting | Aspose.Cells PDF export size analysis | C# get PDF file length | default PDF export size measurement

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, saves it twice as PDF—once with default settings and once with PdfSaveOptions.AllColumnsInOnePagePerSheet enabled—then reads each PDF's byte length and prints the sizes, indicating which export produces a larger or smaller file.
class PdfSizeComparison
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(excelPath);

        // Export PDF with default settings
        string defaultPdfPath = "default.pdf";
        workbook.Save(defaultPdfPath, SaveFormat.Pdf);

        // Export PDF with AllColumnsInOnePagePerSheet option enabled
        string allColumnsPdfPath = "allcolumns.pdf";
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure all columns of each sheet fit on one page
            AllColumnsInOnePagePerSheet = true
        };
        workbook.Save(allColumnsPdfPath, pdfOptions);

        // Get file sizes
        long defaultSize = new FileInfo(defaultPdfPath).Length;
        long allColumnsSize = new FileInfo(allColumnsPdfPath).Length;

        // Output comparison results
        Console.WriteLine($"Default PDF size: {defaultSize} bytes");
        Console.WriteLine($"AllColumnsInOnePagePerSheet PDF size: {allColumnsSize} bytes");

        if (allColumnsSize > defaultSize)
            Console.WriteLine("AllColumnsInOnePagePerSheet PDF is larger.");
        else if (allColumnsSize < defaultSize)
            Console.WriteLine("AllColumnsInOnePagePerSheet PDF is smaller.");
        else
            Console.WriteLine("Both PDFs have the same size.");
    }
}
