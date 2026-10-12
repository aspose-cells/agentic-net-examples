// Title: Export a C# Aspose.Cells workbook that contains a #DIV/0! formula to PDF without raising exceptions
// AI Prompts: Write C# code that builds a workbook, inserts a division‑by‑zero formula, forces calculation, and saves it as a PDF using Aspose.Cells while guaranteeing no exception is thrown. | Demonstrate configuring PdfSaveOptions in Aspose.Cells for .NET to handle cells with formula errors during PDF generation. | Provide a pattern for catching and logging unexpected errors when converting an Excel file with #DIV/0! cells to PDF.
// Common Searches: how to export an Excel file with #DIV/0! cells to PDF using Aspose.Cells .NET without errors | Aspose.Cells PdfSaveOptions ignore formula errors during PDF conversion | C# save workbook to PDF when worksheet contains division by zero errors | prevent exception on PDF export of Excel workbook that has invalid formulas in Aspose.Cells
// Tags: Aspose.Cells PDF export formula errors | C# PdfSaveOptions suppress cell errors | export workbook with #DIV/0! to PDF | handle division by zero during PDF conversion | calculate formulas before PDF save Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a Workbook, adds a division‑by‑zero formula to generate a #DIV/0! error, forces calculation, verifies the error value, and then saves the workbook to a PDF file using PdfSaveOptions while ensuring that no exception is thrown during the export.
class ExportIgnoreErrorsDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with data that will cause a formula error (#DIV/0!)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(0);
            sheet.Cells["B1"].Formula = "=A1/A2"; // Division by zero

            // Force calculation to generate the error in the cell
            workbook.CalculateFormula();

            // Verify that the cell indeed contains an error
            string errorValue = sheet.Cells["B1"].StringValue; // Should be "#DIV/0!"
            Console.WriteLine($"Cell B1 value before export: {errorValue}");

            // Set PDF save options (no IgnoreErrors property in current API)
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Export the workbook to PDF
            workbook.Save("ExportedWithIgnoreErrors.pdf", saveOptions);
            Console.WriteLine("Export completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Operation failed: {ex.Message}");
        }
    }
}
