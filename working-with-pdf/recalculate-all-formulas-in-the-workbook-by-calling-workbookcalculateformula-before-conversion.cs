// Title: Recalculate Excel formulas before PDF conversion with Aspose.Cells for .NET (C#)
// Description: Loads an XLSX workbook, forces a full formula evaluation using Workbook.CalculateFormula(), then exports the file to PDF with PdfSaveOptions.CalculateFormula disabled. Includes file‑existence checks and exception handling.
// Keywords: Aspose.Cells | C# | .NET | Workbook.CalculateFormula | PdfSaveOptions | Excel to PDF conversion | formula recalculation | disable PDF calculation | export Excel as PDF | Aspose.Cells example
// Common Searches: Aspose.Cells recalculate formulas before PDF export | C# calculate all formulas then save as PDF | PdfSaveOptions.CalculateFormula false usage | How to force formula evaluation in Aspose.Cells | Excel to PDF conversion with pre‑calculated formulas
// Developer Intent: Evaluate every formula in a workbook first, then generate a PDF without triggering another calculation pass.
// Use Cases: Creating financial reports where the PDF must show the latest computed values. | Batch‑processing spreadsheets to ensure formulas are up‑to‑date before archiving as PDFs. | Generating printable documents from Excel files while minimizing conversion time by skipping redundant calculations.
// AI Prompts: Generate C# code that opens an .xlsx file, runs Workbook.CalculateFormula(), and saves it as a PDF with PdfSaveOptions.CalculateFormula set to false using Aspose.Cells. | Explain why PdfSaveOptions.CalculateFormula should be disabled after manually recalculating formulas in Aspose.Cells. | Provide best‑practice tips for handling missing input files and runtime exceptions during Excel‑to‑PDF conversion with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PDF conversion if needed

namespace AsposeCellsExamples
{
    // Loads an XLSX workbook, forces a full formula evaluation using Workbook.CalculateFormula(), then exports the file to PDF with PdfSaveOptions.CalculateFormula disabled. Includes file‑existence checks and exception handling.
    public class RecalculateFormulasBeforeConversion
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // Set PDF save options; formulas are already calculated
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    CalculateFormula = false
                };

                // Save the workbook as a PDF file
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine("Formulas recalculated and workbook converted to PDF successfully.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
