using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Sample Data");
        worksheet.Cells["B2"].PutValue(123);

        // Ensure that gridlines are not printed via the worksheet's page setup
        worksheet.PageSetup.PrintGridlines = false;

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // NOTE: The property `RenderSolidGridlines` is not present in the documented
        // Aspose.Cells API for PdfSaveOptions. If your version of the library
        // provides this property, you can disable solid gridlines like this:
        // pdfSaveOptions.RenderSolidGridlines = false;
        // Since the property is undocumented, the line is commented out to keep the code compilable.

        // Save the workbook as a PDF with the specified options
        workbook.Save("CleanLayout.pdf", pdfSaveOptions);
    }
}