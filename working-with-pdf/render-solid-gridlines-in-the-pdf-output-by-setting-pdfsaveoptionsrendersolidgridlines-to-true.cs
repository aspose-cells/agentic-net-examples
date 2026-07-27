using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Author: Aspose.Cells .NET example – renders PDF with gridlines (solid gridlines option may require a newer API)

class Program
{
    static void Main()
    {
        // Create a new workbook (replace with the provided create rule if available)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some data
        sheet.Cells["A1"].PutValue("Sample Data");
        sheet.Cells["B2"].PutValue(123);
        sheet.Cells["C3"].PutValue(DateTime.Now);

        // Ensure gridlines are visible in the worksheet
        sheet.IsGridlinesVisible = true;

        // Create PDF save options (replace with the provided save rule if available)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // The property to render solid gridlines is not documented in the current API version.
        // If your Aspose.Cells version supports it, uncomment the line below:
        // pdfOptions.RenderSolidGridlines = true;

        // Save the workbook as PDF
        workbook.Save("Output.pdf", pdfOptions);
    }
}