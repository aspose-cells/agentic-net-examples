// Title: Apply Black‑and‑White Printing to Confidential Worksheets and Export to PDF with Aspose.Cells for .NET
// Description: Demonstrates how to flag worksheets whose names contain "Confidential" with PageSetup.BlackAndWhite, keep other sheets in color, configure PdfSaveOptions (e.g., ignore blank pages), and save the workbook as a single PDF where confidential data appears in grayscale.
// Keywords: Aspose.Cells black and white worksheet | conditional PDF export .NET | PageSetup.BlackAndWhite property | export confidential sheets to PDF | ignore blank pages Aspose.Cells | grayscale PDF Aspose.Cells | C# workbook PDF conversion | selective worksheet formatting
// Common Searches: how to set black and white mode for specific sheets in Aspose.Cells | export confidential worksheets as grayscale PDF using C# | Aspose.Cells conditional PageSetup settings before PDF save | ignore empty pages when saving workbook to PDF Aspose.Cells
// Developer Intent: Set a grayscale printing mode only on confidential worksheets and generate a PDF that mixes black‑and‑white and color pages.
// Use Cases: Mark worksheets with "Confidential" in the title as black‑and‑white while leaving public sheets in color, then create a combined PDF. | Load an existing workbook, apply PageSetup.BlackAndWhite based on a custom list of confidential sheet names, and export with PdfSaveOptions.PrintingPageType = IgnoreBlank. | Generate reports where sensitive data is visually de‑emphasized (grayscale) without affecting the visual style of non‑confidential sections.
// AI Prompts: Write C# code using Aspose.Cells to set PageSetup.BlackAndWhite = true for worksheets whose name contains a confidentiality keyword and save the workbook as a PDF. | Show how to configure PdfSaveOptions to ignore blank pages while exporting a workbook that contains both color and grayscale worksheets. | Explain how to read a list of confidential worksheet names from an external source, apply the black‑and‑white setting, and export the result to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Demonstrates how to flag worksheets whose names contain "Confidential" with PageSetup.BlackAndWhite, keep other sheets in color, configure PdfSaveOptions (e.g., ignore blank pages), and save the workbook as a single PDF where confidential data appears in grayscale.
    public class ConfidentialBlackAndWhitePdfExport
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Sample data: add three worksheets, two of them are confidential
                // -------------------------------------------------
                Worksheet ws1 = workbook.Worksheets[0];
                ws1.Name = "Summary";
                ws1.Cells["A1"].PutValue("Public Data");

                Worksheet wsConf1 = workbook.Worksheets.Add("Confidential_Q1");
                wsConf1.Cells["A1"].PutValue("Secret Data Q1");

                Worksheet wsConf2 = workbook.Worksheets.Add("Confidential_Q2");
                wsConf2.Cells["A1"].PutValue("Secret Data Q2");

                // -------------------------------------------------
                // Apply black‑and‑white printing mode only to confidential sheets
                // -------------------------------------------------
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Define your own criteria for confidentiality.
                    // Here we treat any sheet whose name contains "Confidential" as confidential.
                    if (sheet.Name.IndexOf("Confidential", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Set the BlackAndWhite property to true for printing.
                        sheet.PageSetup.BlackAndWhite = true;
                    }
                    else
                    {
                        // Ensure non‑confidential sheets retain their original color mode.
                        sheet.PageSetup.BlackAndWhite = false;
                    }
                }

                // -------------------------------------------------
                // Prepare PDF save options (you can customize further if needed)
                // -------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Example: ignore blank pages to keep the PDF tidy
                    PrintingPageType = PrintingPageType.IgnoreBlank
                };

                // -------------------------------------------------
                // Export the workbook to PDF. Only the confidential sheets will be rendered in black‑and‑white.
                // -------------------------------------------------
                string outputPath = "ConfidentialReport.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook exported to PDF at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
