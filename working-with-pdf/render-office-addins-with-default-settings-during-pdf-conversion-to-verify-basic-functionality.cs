// Title: C# – Render Office Add‑In (WebExtension) in PDF with Aspose.Cells default settings
// Description: Creates a workbook, adds a WebExtension with only the required Id and StoreName, applies the default PdfSaveOptions, and saves the file as PDF to confirm that the Office Add‑In is rendered correctly without extra configuration.
// Keywords: Aspose.Cells | C# | WebExtension | Office Add‑In | PDF conversion | default PdfSaveOptions | .NET | render add‑in in PDF | save workbook as PDF
// Common Searches: Aspose.Cells render WebExtension in PDF | C# add Office Add‑In to workbook and save as PDF | default PDF conversion settings for WebExtension | how to include Office Add‑In in PDF output using Aspose.Cells | verify WebExtension appears in generated PDF
// Developer Intent: Confirm that an Office Add‑In (WebExtension) is rendered correctly when a workbook is saved to PDF using Aspose.Cells with its default PDF options.
// Use Cases: Quick validation that a newly added WebExtension is visible in the PDF output. | Automated regression test to ensure default PDF settings preserve add‑in rendering. | Generating sample PDFs for documentation that showcase embedded Office Add‑Ins.
// AI Prompts: Show C# code to add a WebExtension with custom properties and save the workbook to PDF using specific PdfSaveOptions. | Explain how to programmatically verify that the Office Add‑In appears in the generated PDF with Aspose.Cells. | Demonstrate how to enable or disable WebExtension rendering during PDF conversion in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.WebExtensions;
using Aspose.Cells.Rendering;

// Creates a workbook, adds a WebExtension with only the required Id and StoreName, applies the default PdfSaveOptions, and saves the file as PDF to confirm that the Office Add‑In is rendered correctly without extra configuration.
class Program
{
    static void Main()
    {
        // Create a new workbook (default constructor)
        Workbook workbook = new Workbook();

        // Add an Office Add‑in (WebExtension) with default settings
        WebExtensionCollection webExtensions = workbook.Worksheets.WebExtensions;
        int extensionIndex = webExtensions.Add();               // creates a new WebExtension
        WebExtension webExtension = webExtensions[extensionIndex];

        // Minimal reference information (required fields)
        webExtension.Reference.Id = "exampleAddIn";
        webExtension.Reference.StoreName = "ExampleStore";

        // No additional configuration – use default PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();       // default settings

        // Save the workbook as PDF to verify that the add‑in is rendered correctly
        workbook.Save("AddInDefaultPdf.pdf", pdfOptions);
    }
}
