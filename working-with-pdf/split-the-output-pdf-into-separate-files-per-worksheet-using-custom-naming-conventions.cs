// Title: C# – Split Excel Workbook into Separate PDFs per Worksheet with Custom Naming (Aspose.Cells)
// Description: A C# example that loads or creates an Excel workbook, loops through its worksheets, and saves each one as an individual PDF using Aspose.Cells. PdfSaveOptions.SheetSet isolates the current sheet, OnePagePerSheet forces a single‑page PDF, and a dynamic filename such as "Sheet_1_Sheet1.pdf" is generated for every output file.
// Keywords: Aspose.Cells | C# | .NET | PdfSaveOptions | SheetSet | OnePagePerSheet | export worksheet to PDF | split workbook PDF | custom PDF filename | Excel to PDF per sheet | programmatic PDF generation
// Common Searches: export each Excel sheet to separate PDF Aspose.Cells | Aspose.Cells PDFSaveOptions SheetSet example | C# split workbook into multiple PDFs | custom file name for each PDF Aspose.Cells | OnePagePerSheet option .NET
// Developer Intent: Generate an individual PDF file for every worksheet in a workbook, applying a custom naming convention.
// Use Cases: Distribute department‑specific sheets of a financial report as separate PDFs. | Automate per‑sheet PDF creation for client‑tailored Excel workbooks. | Produce single‑page PDFs for each worksheet to embed in web portals or document management systems.
// AI Prompts: Show how to add a timestamp to the PDF file name in this Aspose.Cells example. | Explain how to combine several worksheets into one PDF using PdfSaveOptions.SheetSet. | Provide code to save each worksheet as a password‑protected PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfSplitDemo
{
    // A C# example that loads or creates an Excel workbook, loops through its worksheets, and saves each one as an individual PDF using Aspose.Cells. PdfSaveOptions.SheetSet isolates the current sheet, OnePagePerSheet forces a single‑page PDF, and a dynamic filename such as "Sheet_1_Sheet1.pdf" is generated for every output file.
    class Program
    {
        static void Main()
        {
            // Load or create a workbook
            Workbook workbook = new Workbook(); // create a new workbook

            // Add sample worksheets with data
            for (int i = 0; i < 3; i++)
            {
                Worksheet sheet = i == 0 ? workbook.Worksheets[0] : workbook.Worksheets.Add($"Sheet{i + 1}");
                sheet.Cells["A1"].PutValue($"Data for {sheet.Name}");
                sheet.Cells["A2"].PutValue(DateTime.Now);
            }

            // Iterate through each worksheet and save it as an individual PDF
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Prepare PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Configure the SheetSet to include only the current worksheet (zero‑based index)
                pdfOptions.SheetSet = new SheetSet(new int[] { i });

                // Optional: ensure each sheet is rendered on a single page
                pdfOptions.OnePagePerSheet = true;

                // Build a custom file name: "Sheet_{index}_{name}.pdf"
                string fileName = $"Sheet_{i + 1}_{workbook.Worksheets[i].Name}.pdf";

                // Save the workbook as PDF using the configured options
                workbook.Save(fileName, pdfOptions);
            }
        }
    }
}
