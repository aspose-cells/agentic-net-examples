using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set explicit column widths to demonstrate that they are preserved
        sheet.Cells.SetColumnWidth(0, 20); // Column A width
        sheet.Cells.SetColumnWidth(1, 30); // Column B width

        // Populate sample cells
        sheet.Cells["A1"].PutValue("This is a long text that requires a wide column");
        sheet.Cells["B1"].PutValue(12345);

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Minimize the PDF file size
            OptimizationType = PdfOptimizationType.MinimumSize,
            // Preserve original column widths (do not force all columns onto one page)
            AllColumnsInOnePagePerSheet = false,
            OnePagePerSheet = false
        };

        // Save the workbook as a PDF with the specified options
        workbook.Save("PreservedColumns.pdf", pdfOptions);
    }
}

// Author: Aspose.Cells .NET example – sets MinimumSize optimization while keeping column widths.