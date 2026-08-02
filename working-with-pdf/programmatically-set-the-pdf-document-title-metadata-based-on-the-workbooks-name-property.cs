using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – sets PDF title from workbook name
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Assign a file name to the workbook (used as the title)
        workbook.FileName = "SampleWorkbook.xlsx";

        // Set the built‑in document title to match the workbook's file name
        workbook.BuiltInDocumentProperties.Title = workbook.FileName;

        // Configure PDF save options to display the document title in the viewer's title bar
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DisplayDocTitle = true
        };

        // Save the workbook as PDF; the PDF will carry the title metadata
        workbook.Save("SampleWorkbook.pdf", pdfOptions);
    }
}