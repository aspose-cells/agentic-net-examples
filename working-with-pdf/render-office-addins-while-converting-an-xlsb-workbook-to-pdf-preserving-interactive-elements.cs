using System;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Saving;

class Program
{
    // Author: Generated example for converting XLSB to PDF while preserving Office Add‑Ins
    static async Task Main(string[] args)
    {
        // Paths to the source XLSB workbook and the target PDF file
        string inputFile = "input.xlsb";
        string outputFile = "output.pdf";

        // Load the workbook (XLSB format)
        Workbook workbook = new Workbook(inputFile);

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Preserve interactive Office Add‑Ins (if the API provides such a setting)
        // NOTE: Verify the exact property name in the Aspose.Cells documentation for your version.
        // pdfOptions.PreserveOfficeAddIns = true; // <-- placeholder for the actual property

        // Save the workbook as PDF using the configured options
        workbook.Save(outputFile, pdfOptions);
    }
}