using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some data
        worksheet.Cells["A1"].PutValue("Column 1");
        worksheet.Cells["B1"].PutValue("Column 2");
        worksheet.Cells["A2"].PutValue("Data 1");
        worksheet.Cells["B2"].PutValue("Data 2");

        // Set explicit column widths to ensure they are preserved in the PDF
        worksheet.Cells.SetColumnWidth(0, 25); // Width for column A
        worksheet.Cells.SetColumnWidth(1, 30); // Width for column B

        // Create PDF save options and set the optimization type to MinimumSize
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF while preserving the original column widths
        workbook.Save("PreservedColumns_MinSize.pdf", pdfSaveOptions);
    }
}