using System;
using System.IO;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – detects empty worksheets and skips them when saving to PDF
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Hide worksheets that are empty to prevent blank pages in the PDF
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Since Aspose.Cells 21.5.2, MaxDisplayRange is null for an empty sheet
            if (sheet.Cells.MaxDisplayRange == null)
            {
                sheet.IsVisible = false; // Skip this sheet during PDF rendering
            }
        }

        // Configure PDF save options to ignore blank pages and not output a blank page when nothing is printed
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PrintingPageType = PrintingPageType.IgnoreBlank,
            OutputBlankPageWhenNothingToPrint = false
        };

        // Save the workbook as PDF
        string outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Result.pdf");
        workbook.Save(outputPath, pdfOptions);
    }
}