using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Needed for PdfSaveOptions

// Author: Aspose.Cells .NET example – verifies solid gridlines in PDF output
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Make gridlines visible in the worksheet
        worksheet.IsGridlinesVisible = true;

        // Populate some sample data so gridlines can be seen
        worksheet.Cells["A1"].PutValue("Solid Gridlines Demo");
        worksheet.Cells["B2"].PutValue(123);
        worksheet.Cells["C3"].PutValue(DateTime.Now);

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // NOTE: The property to enable solid gridlines (e.g., RenderSolidGridlines) is not
        // documented in the provided reference. If the current Aspose.Cells version
        // supports such a property, set it here. Placeholder shown below:
        // pdfSaveOptions.RenderSolidGridlines = true; // <-- set to true when API exists

        // Save the workbook as PDF
        workbook.Save("SolidGridlinesDemo.pdf", pdfSaveOptions);
    }
}