// Title: Apply Black‑and‑White Printing to Confidential Worksheets and Export to PDF with Aspose.Cells for .NET (C#)
// Description: Loads a workbook, detects worksheets whose names contain "Confidential", sets their PageSetup.BlackAndWhite flag, and saves the file as a PDF using PdfSaveOptions (visible sheets only). Confidential sheets are rendered in grayscale while other sheets keep their original colors.
// Keywords: Aspose.Cells black and white PDF | PageSetup.BlackAndWhite C# | selective grayscale Aspose.Cells | export confidential worksheets to PDF | PdfSaveOptions SheetSet Visible | C# Aspose.Cells PDF export
// Common Searches: Aspose.Cells set black and white for specific sheets | C# export only visible worksheets to PDF | how to grayscale confidential worksheets in Aspose.Cells | apply PageSetup.BlackAndWhite before PDF conversion | selective grayscale export Aspose.Cells .NET
// Developer Intent: Mark worksheets that contain confidential data as black‑and‑white and generate a PDF that includes only visible sheets.
// Use Cases: Create compliance‑ready PDFs where sensitive worksheets are shown in grayscale while the rest remain in color. | Automate batch processing of workbooks to enforce a naming‑based grayscale rule before distribution. | Generate reports that hide hidden sheets and protect confidential content by rendering it in black‑and‑white.
// AI Prompts: Write C# code with Aspose.Cells to enable PageSetup.BlackAndWhite for worksheets whose name includes "Confidential" and save the workbook as a PDF containing only visible sheets. | Explain the effect of the PageSetup.BlackAndWhite property on PDF rendering in Aspose.Cells and any scenarios where it does not apply. | Suggest an alternative method to apply grayscale to selected worksheets without altering the color settings of other sheets during PDF export.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsBlackAndWhiteDemo
{
    // Loads a workbook, detects worksheets whose names contain "Confidential", sets their PageSetup.BlackAndWhite flag, and saves the file as a PDF using PdfSaveOptions (visible sheets only). Confidential sheets are rendered in grayscale while other sheets keep their original colors.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Identify confidential sheets (e.g., name contains "Confidential")
                if (sheet.Name.IndexOf("Confidential", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Enable black‑and‑white printing for this sheet
                    sheet.PageSetup.BlackAndWhite = true;
                }
            }

            // Configure PDF save options (you can adjust other options as needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: export only visible sheets (optional)
                SheetSet = SheetSet.Visible
            };

            // Save the workbook to PDF
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved to PDF with black‑and‑white applied to confidential sheets: {outputPath}");
        }
    }
}
