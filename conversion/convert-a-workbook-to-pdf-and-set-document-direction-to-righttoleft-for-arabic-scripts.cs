// Title: C# – Convert Aspose.Cells Workbook to PDF with Right‑to‑Left Layout for Arabic
// Description: Demonstrates how to create a workbook, enable the DisplayRightToLeft property for Arabic scripts, insert sample Arabic text, and save the sheet as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PDF conversion C# | right‑to‑left Excel PDF | Arabic RTL PDF Aspose | DisplayRightToLeft property | .NET workbook to PDF | Arabic Excel export | C# Aspose.Cells example
// Common Searches: Aspose.Cells convert Excel to PDF RTL Arabic | C# set worksheet right‑to‑left before PDF export | DisplayRightToLeft Aspose.Cells PDF output | Generate Arabic PDF from Excel using .NET | How to export Arabic sheet as PDF with Aspose
// Developer Intent: Create a PDF from an Excel workbook while forcing right‑to‑left rendering for Arabic content.
// Use Cases: Produce Arabic reports or invoices in PDF with correct RTL orientation. | Automate multilingual document generation where Arabic sheets require RTL layout. | Integrate PDF export into web portals serving Middle‑East users.
// AI Prompts: Write C# code that opens an existing Excel file, sets DisplayRightToLeft on a specific worksheet, and saves it as a PDF with Aspose.Cells. | Explain whether additional font embedding is needed when exporting Arabic text to PDF with Aspose.Cells. | Provide a script to batch‑process a folder of workbooks, applying RTL layout to each sheet and exporting PDFs.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, enable the DisplayRightToLeft property for Arabic scripts, insert sample Arabic text, and save the sheet as a PDF using Aspose.Cells for .NET.
class ConvertWorkbookToPdfRtl
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable right‑to‑left display for Arabic scripts
        worksheet.DisplayRightToLeft = true;

        // Add sample Arabic text
        worksheet.Cells["A1"].PutValue("مرحبا بالعالم"); // "Hello World" in Arabic

        // Save the workbook as PDF
        workbook.Save("ArabicRtlOutput.pdf", SaveFormat.Pdf);
    }
}
