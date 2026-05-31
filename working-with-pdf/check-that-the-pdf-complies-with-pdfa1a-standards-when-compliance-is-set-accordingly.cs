using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfA1aComplianceDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("PDF/A-1a compliance test");

        // Create PDF save options and set compliance to PDF/A-1a
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1a
        };

        // Save the workbook as a PDF file with the specified compliance level
        string outputFile = "PdfA1aOutput.pdf";
        workbook.Save(outputFile, saveOptions);

        Console.WriteLine($"PDF saved with PDF/A-1a compliance: {outputFile}");
    }
}