// Title: Load an XLS workbook, recalculate all formulas, and export to PDF (owner password not supported) with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xls file using Aspose.Cells, runs CalculateFormula, and saves the workbook as a PDF. | Show how to verify the source Excel file exists before converting it to PDF with Aspose.Cells in C#. | Explain how to apply an owner password to a PDF generated from Excel using Aspose.Cells and why older versions lack this feature. | Provide a try‑catch pattern for handling errors during Excel‑to‑PDF conversion with formula evaluation in Aspose.Cells.
// Common Searches: asp.net calculate formulas in xls before converting to pdf with aspose.cells | c# aspose.cells export excel to pdf with owner password | how to set pdf encryption when saving workbook as pdf using aspose.cells | aspose.cells pdf export formula evaluation example c#
// Tags: calculate-formulas-aspocells | xls-to-pdf-aspocells | pdf-encryption-aspocells | workbook-save-pdf-aspocells | aspose-cells-formula-evaluation

using System;
using System.IO;
using Aspose.Cells;

// The example checks that input.xls exists, loads it into an Aspose.Cells Workbook, forces recalculation of all worksheet formulas, and saves the result as output.pdf. Because the Aspose.Cells version used does not expose PDF password protection, the PDF is created without encryption; any runtime errors are caught and logged to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the XLS file
            Workbook workbook = new Workbook(inputPath);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook as a PDF (no encryption due to API version constraints)
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
