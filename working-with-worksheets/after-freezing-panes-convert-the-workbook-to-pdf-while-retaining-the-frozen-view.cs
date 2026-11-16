// Title: How to freeze panes in an Aspose.Cells worksheet and export the workbook to PDF while keeping the frozen view (C#)
// Common Searches: C# Aspose.Cells keep frozen rows and columns when saving as PDF | Export Excel workbook to PDF with freeze panes using Aspose.Cells in .NET | Aspose.Cells PDF export preserving frozen panes in C# application | Save workbook with FreezePanes to PDF using PdfSaveOptions in Aspose.Cells | Maintain frozen pane layout after converting Excel to PDF with Aspose.Cells
// Tags: Aspose.Cells FreezePanes PDF export | PdfSaveOptions preserve frozen view | C# export Excel to PDF with frozen panes | Workbook.Save with PdfSaveOptions Aspose.Cells | maintain worksheet freeze state PDF conversion

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, fills it with data, freezes the first row and column via FreezePanes, configures PdfSaveOptions (OnePagePerSheet = false) to retain the layout, ensures the output directory exists, and saves the workbook as a PDF while preserving the frozen pane view.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            for (int i = 0; i < 100; i++)
            {
                sheet.Cells[i + 1, 0].PutValue($"Row {i + 1}");
                sheet.Cells[i + 1, 1].PutValue(i * 10);
            }

            // Freeze the first row and first column (A1 cell)
            // Parameters: (firstRow, firstColumn, totalRows, totalColumns)
            // This freezes rows above row 2 and columns left of column B.
            sheet.FreezePanes(1, 1, 1, 1);

            // Prepare PDF save options (retain frozen view automatically)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep the original layout; frozen panes are reflected in the PDF.
                OnePagePerSheet = false
            };

            string outputPath = "output.pdf";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF (lifecycle rule: save)
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
