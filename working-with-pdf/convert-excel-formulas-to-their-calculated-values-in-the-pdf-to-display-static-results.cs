// Title: Export Excel to PDF with Calculated Formulas as Static Values using Aspose.Cells for .NET
// Description: Shows how to create a workbook, insert numbers, apply a SUM formula, enable PdfSaveOptions.CalculateFormula, and save the file as a PDF where the cell displays the computed result (30) instead of the formula text.
// Keywords: Aspose.Cells | PDF export | calculate formulas | static PDF | C# | .NET | PdfSaveOptions | Excel to PDF | formula evaluation | workbook conversion
// Common Searches: Aspose.Cells export Excel to PDF with calculated values | PdfSaveOptions.CalculateFormula example C# | how to hide formulas in PDF generated from Excel | convert Excel workbook to static PDF using Aspose | evaluate Excel formulas before PDF export .NET
// Developer Intent: Generate a PDF from an Excel workbook where all formulas are evaluated and rendered as fixed values.
// Use Cases: Produce client‑ready PDF reports that show totals without exposing underlying formulas. | Automate financial statement distribution as static PDFs from dynamic spreadsheets. | Create printable PDFs from web‑based worksheets where calculations must be pre‑computed.
// AI Prompts: Provide C# code to export an Aspose.Cells workbook to PDF with formulas evaluated. | Explain the effect of PdfSaveOptions.CalculateFormula on PDF output and any constraints. | Show how to load an existing Excel file, recalculate all formulas, and save it as a static PDF using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfFormulaDemo
{
    // Shows how to create a workbook, insert numbers, apply a SUM formula, enable PdfSaveOptions.CalculateFormula, and save the file as a PDF where the cell displays the computed result (30) instead of the formula text.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            // Set a formula that will be calculated
            cells["A3"].Formula = "=SUM(A1:A2)";

            // Create PDF save options and enable formula calculation
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // This ensures that formulas are evaluated before the PDF is generated
                CalculateFormula = true
            };

            // Save the workbook as PDF with calculated values
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.pdf");
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved to: {outputPath}");
            // The PDF will display the static result (30) instead of the formula text.
        }
    }
}
