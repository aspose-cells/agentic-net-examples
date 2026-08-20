// Title: Convert Excel to PDF/A‑1b with Aspose.Cells in C# (.NET)
// Description: Shows how to create a workbook, add sample data, set PdfSaveOptions.Compliance to PdfCompliance.PdfA1b, and save the file as a PDF/A‑1b document for long‑term archival using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | PDF/A-1b | PdfSaveOptions | PdfCompliance | Excel to PDF/A | archival PDF | document compliance | long‑term storage
// Common Searches: Aspose.Cells PDF/A-1b export C# | How to save workbook as PDF/A using Aspose.Cells .NET | Set PDF compliance level Aspose.Cells | Generate archival PDF from Excel Aspose | PdfSaveOptions Compliance property example
// Developer Intent: Save an Excel workbook as a PDF/A‑1b file by configuring the compliance option.
// Use Cases: Create regulatory‑compliant financial reports in PDF/A‑1b. | Archive generated spreadsheets for legal retention periods. | Integrate PDF/A conversion into automated reporting pipelines. | Batch process multiple workbooks to PDF/A‑1b for document management systems.
// AI Prompts: Show me how to change the code to use PDF/A‑2b compliance instead of PDF/A‑1b. | Generate a reusable method that takes a Workbook, output path, and optional metadata, then saves the workbook as PDF/A‑1b. | Write a script that loops through a folder of .xlsx files and converts each to PDF/A‑1b using Aspose.Cells. | Explain how to embed XMP metadata for PDF/A compliance when saving with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to create a workbook, add sample data, set PdfSaveOptions.Compliance to PdfCompliance.PdfA1b, and save the file as a PDF/A‑1b document for long‑term archival using Aspose.Cells for .NET.
class PdfAComplianceDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet and add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Aspose.Cells PDF/A compliance example");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Create PDF save options and set the compliance level to PDF/A-1b (rule: PdfSaveOptions.Compliance)
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the workbook as a PDF with the specified compliance level (lifecycle rule: save)
        workbook.Save("PdfA1b_Output.pdf", saveOptions);
    }
}
