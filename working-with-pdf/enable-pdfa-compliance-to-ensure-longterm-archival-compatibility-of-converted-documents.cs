using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample content
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF/A compliant document");

        // Configure PDF save options for PDF/A-1b compliance
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.Compliance = PdfCompliance.PdfA1b; // Enable PDF/A-1b compliance

        // Save the workbook as a PDF file with the specified compliance level
        workbook.Save("output.pdf", saveOptions);
    }
}

// Author: Aspose.Cells .NET example for PDF/A compliance.