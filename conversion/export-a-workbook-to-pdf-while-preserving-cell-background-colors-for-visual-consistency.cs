// Title: How to export an Aspose.Cells workbook to PDF with cell background colors retained using C#
// AI Prompts: Generate C# code that creates a workbook, applies solid fill colors to cells, and saves it as a PDF using Aspose.Cells PdfSaveOptions so the colors stay visible. | Show the required PdfSaveOptions configuration in C# to prevent background shading from being lost during Excel‑to‑PDF conversion with Aspose.Cells. | Explain which Aspose.Cells properties must be enabled to keep cell fill colors when exporting a worksheet to PDF in a .NET application.
// Common Searches: Aspose.Cells C# export Excel to PDF keep cell fill colors | PdfSaveOptions preserve background shading Aspose.Cells .NET | How to retain worksheet cell colors when saving as PDF with Aspose.Cells | C# code sample for exporting a colored Excel sheet to PDF using Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions background shading | export Excel to PDF with cell colors C# | preserve cell fill colors Aspose.Cells PDF conversion | C# Aspose.Cells PDF export settings

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // The example creates a workbook, fills cells with data, applies solid background colors to header and data rows, calculates formulas, configures PdfSaveOptions with ExportDocumentStructure enabled, and saves the workbook as a PDF while preserving all cell background colors.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apples");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Bananas");
                sheet.Cells["B3"].PutValue(85);

                // Apply background colors to header row
                Style styleHeader = workbook.CreateStyle();
                styleHeader.ForegroundColor = Color.LightGray;
                styleHeader.Pattern = BackgroundType.Solid;
                sheet.Cells.CreateRange("A1:B1").SetStyle(styleHeader);

                // Apply background colors to Apples row
                Style styleApples = workbook.CreateStyle();
                styleApples.ForegroundColor = Color.LightGreen;
                styleApples.Pattern = BackgroundType.Solid;
                sheet.Cells.CreateRange("A2:B2").SetStyle(styleApples);

                // Apply background colors to Bananas row
                Style styleBananas = workbook.CreateStyle();
                styleBananas.ForegroundColor = Color.LightYellow;
                styleBananas.Pattern = BackgroundType.Solid;
                sheet.Cells.CreateRange("A3:B3").SetStyle(styleBananas);

                // Ensure any formulas are calculated before saving
                workbook.CalculateFormula();

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Preserve the document structure (rule: ExportDocumentStructure property)
                    ExportDocumentStructure = true
                    // Optionally, preserve gridline color if needed
                    // GridlineColor = Color.DarkGray
                };

                // Save the workbook to PDF while preserving cell background colors (lifecycle rule: save)
                string outputPath = "WorkbookWithColors.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine("PDF exported successfully with cell background colors preserved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
