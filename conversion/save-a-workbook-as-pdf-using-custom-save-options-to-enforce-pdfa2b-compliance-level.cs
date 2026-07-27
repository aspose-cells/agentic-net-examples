// Title: Create PDF/A‑2b compliant PDF from an Aspose.Cells workbook using C# PdfSaveOptions
// Description: Shows how to build a Workbook, add sample data, set PdfSaveOptions.Compliance to PdfCompliance.PdfA2b, and save the file as a PDF/A‑2b‑compatible document (PdfA2bOutput.pdf) with Aspose.Cells for .NET.
// Keywords: Aspose.Cells PDF/A-2b | C# PdfSaveOptions | PDF/A-2b export .NET | PdfCompliance.PdfA2b | Excel to PDF/A-2b conversion | archival PDF generation | ISO 19005-2 compliance | Aspose.Cells PDF export | C# workbook to PDF/A-2b
// Common Searches: Aspose.Cells export to PDF/A-2b C# | PdfSaveOptions compliance PDF/A-2b example | How to set PDF/A-2b in Aspose.Cells | C# generate archival PDF from Excel | PDF/A-2b conversion using Aspose.Cells .NET
// Developer Intent: Generate a PDF that meets the PDF/A‑2b archival standard directly from an Aspose.Cells workbook in a C# application.
// Use Cases: Long‑term preservation of financial or regulatory reports as PDF/A‑2b files. | Ensuring exported spreadsheets comply with document‑management policies that require ISO‑19005‑2. | Creating PDF/A‑2b documents for electronic filing systems that mandate archival‑grade PDFs.
// AI Prompts: Provide C# code that saves an Aspose.Cells workbook as PDF/A‑2b with custom page size and orientation. | Explain how to programmatically verify that the generated PDF conforms to PDF/A‑2b standards using Aspose.PDF or third‑party validators. | Show a .NET script to batch convert a folder of Excel files to PDF/A‑2b using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to build a Workbook, add sample data, set PdfSaveOptions.Compliance to PdfCompliance.PdfA2b, and save the file as a PDF/A‑2b‑compatible document (PdfA2bOutput.pdf) with Aspose.Cells for .NET.
class PdfA2bSaveDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF/A-2b compliance demo");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Create PDF save options and set the compliance level to PDF/A-2b
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.Compliance = PdfCompliance.PdfA2b;

        // Save the workbook as a PDF file using the specified options
        workbook.Save("PdfA2bOutput.pdf", saveOptions);
    }
}
