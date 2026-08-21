// Title: Recalculate Formulas and Export to PDF with Aspose.Cells (C#)
// Description: Demonstrates how to trigger Workbook.CalculateFormula to update every cell, optionally enable PdfSaveOptions.CalculateFormula, and save the workbook as a PDF file. The example adds sample data, applies a SUM formula, forces recalculation, and generates Output.pdf.
// Keywords: Aspose.Cells C# PDF export | Workbook.CalculateFormula | PdfSaveOptions CalculateFormula true | recalculate Excel formulas before PDF conversion | .NET Excel to PDF automation | Aspose.Cells formula evaluation | export workbook as PDF with updated formulas | C# Aspose.Cells example
// Common Searches: How to force formula calculation in Aspose.Cells before saving PDF | Aspose.Cells C# recalculate all formulas then export to PDF | PdfSaveOptions.CalculateFormula usage in .NET | Export Excel workbook to PDF with latest formula results | Aspose.Cells workbook.CalculateFormula example
// Developer Intent: Update every formula in a workbook and then generate a PDF that reflects the new values.
// Use Cases: Creating financial statements as PDFs where totals must reflect the latest data. | Automating invoice generation with SUM, AVERAGE, or custom formulas evaluated before PDF output. | Batch converting multiple Excel reports to PDF while guaranteeing all calculations are current.
// AI Prompts: Generate C# code that opens an existing Excel file, runs Workbook.CalculateFormula, disables PdfSaveOptions.CalculateFormula, and saves the result as a PDF. | Show how to set custom page margins in PdfSaveOptions after recalculating formulas with Aspose.Cells. | Write a script that iterates through a folder of .xlsx files, recalculates each workbook's formulas, and exports each one to a PDF with the same filename.

using System;
using Aspose.Cells;

// Demonstrates how to trigger Workbook.CalculateFormula to update every cell, optionally enable PdfSaveOptions.CalculateFormula, and save the workbook as a PDF file. The example adds sample data, applies a SUM formula, forces recalculation, and generates Output.pdf.
public class RecalculateAndExportPdf
{
    // Entry point required for console application
    public static void Main()
    {
        try
        {
            Run();
            Console.WriteLine("PDF exported successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add sample data and a formula
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["A3"].Formula = "=SUM(A1:A2)";

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Set PDF save options (optional: also calculate formulas during save)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            CalculateFormula = true
        };

        // Export the workbook to PDF
        workbook.Save("Output.pdf", pdfOptions);
    }
}
