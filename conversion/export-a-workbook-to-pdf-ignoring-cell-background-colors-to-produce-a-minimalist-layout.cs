// Title: Export Aspose.Cells Workbook to PDF without Cell Background Colors (C#)
// Description: Demonstrates how to generate a minimalist PDF from an Aspose.Cells workbook by applying colors to cells, enabling the BlackAndWhite page‑setup option, and saving with PdfSaveOptions. The resulting PDF contains only text and borders, eliminating all cell fill shading for clean, black‑and‑white printing.
// Keywords: Aspose.Cells PDF export C# | ignore cell background Aspose.Cells | black and white PDF Aspose.Cells | PdfSaveOptions minimal PDF | Worksheet.PageSetup.BlackAndWhite | convert colored sheet to PDF | remove cell fill shading PDF | minimalist PDF layout .NET | Aspose.Cells export without colors | C# generate PDF without background
// Common Searches: Aspose.Cells export PDF without background colors | C# create black and white PDF from Excel | How to ignore cell fill when converting to PDF using Aspose.Cells | Minimalist PDF layout Aspose.Cells .NET | PageSetup.BlackAndWhite PDF conversion example
// Developer Intent: Produce a PDF from a workbook that suppresses all cell background colors for a clean, monochrome output.
// Use Cases: Print price lists or catalogs in economical black‑and‑white format. | Generate reports where colored shading would distract from the data. | Create invoices or statements that require only text and borders, no cell fills.
// AI Prompts: Show C# code to export an Aspose.Cells workbook to PDF while omitting cell background colors. | Explain how PageSetup.BlackAndWhite and PdfSaveOptions work together to create a minimalist PDF. | Provide an example of converting a colored worksheet to a black‑and‑white PDF for printing with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsPdfMinimalist
{
    // Demonstrates how to generate a minimalist PDF from an Aspose.Cells workbook by applying colors to cells, enabling the BlackAndWhite page‑setup option, and saving with PdfSaveOptions. The resulting PDF contains only text and borders, eliminating all cell fill shading for clean, black‑and‑white printing.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (create rule)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.20);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(0.80);

                // Apply background colors to cells (these will be ignored in the PDF)
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.LightBlue;
                style.Pattern = BackgroundType.Solid;
                // Header row style
                sheet.Cells.CreateRange("A1:B1").SetStyle(style);

                style.ForegroundColor = Color.LightYellow;
                // Data rows style
                sheet.Cells.CreateRange("A2:B3").SetStyle(style);

                // Set the worksheet to print in black and white.
                // This causes background colors to be omitted in the PDF output.
                sheet.PageSetup.BlackAndWhite = true;

                // Create PDF save options (no special options needed for this scenario)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook to PDF (save rule)
                string outputPath = "MinimalistLayout.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook saved to PDF without background colors: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
