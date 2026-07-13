using System;
using Aspose.Cells;
using Aspose.Cells.Saving; // PdfSaveOptions resides here

// Author: Generated example for converting XLS to PDF while attempting to preserve interactive controls
class Program
{
    static void Main()
    {
        // Input XLS workbook (can contain form controls, ActiveX, etc.)
        string inputFile = "input.xls";

        // Desired PDF output file
        string outputFile = "output.pdf";

        // Load the workbook from the XLS file
        Workbook workbook = new Workbook(inputFile);

        // Configure PDF save options – aim to keep interactive controls (e.g., form fields)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // NOTE: The exact property to preserve controls may differ between Aspose.Cells versions.
        // If a property such as PreserveFormFields exists, enable it here.
        // pdfOptions.PreserveFormFields = true; // <-- verify against the current API documentation

        // Save the workbook as PDF using the configured options
        workbook.Save(outputFile, pdfOptions);
    }
}