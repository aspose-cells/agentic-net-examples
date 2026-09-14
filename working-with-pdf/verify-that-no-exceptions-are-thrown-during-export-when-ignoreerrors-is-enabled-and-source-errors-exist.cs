// Title: Export an Excel workbook containing #DIV/0! formulas to PDF in C# with Aspose.Cells while suppressing errors
// AI Prompts: Write C# code that creates a worksheet, inserts a formula that results in a #DIV/0! error, verifies the error cell, and saves the workbook to PDF using PdfSaveOptions without raising exceptions. | Show how to configure Aspose.Cells PdfSaveOptions to ignore calculation errors during PDF conversion and produce a PDF file even when cells contain invalid formulas. | Adapt the example to log the address and value of cells that contain errors before exporting the workbook to PDF, ensuring the export proceeds without interruption.
// Common Searches: Aspose.Cells C# export to PDF ignore #DIV/0! errors | How to save Excel workbook with formula errors as PDF using Aspose.Cells | PdfSaveOptions suppress calculation errors during PDF conversion Aspose.Cells | Export Excel file with invalid formulas to PDF without exception in .NET | C# Aspose.Cells ignore errors when converting workbook to PDF
// Tags: Aspose.Cells PDF export error suppression | C# PdfSaveOptions calculation error handling | Aspose.Cells ignore formula errors during PDF conversion | Export workbook containing #DIV/0! to PDF | Aspose.Cells error‑tolerant PDF generation

using System;
using Aspose.Cells;

// The sample creates a workbook, places a formula that triggers a #DIV/0! error, confirms the error cell, and then saves the workbook as a PDF using PdfSaveOptions. The code demonstrates how to suppress calculation errors so the export completes without throwing exceptions.
class ExportIgnoreErrorsDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a formula that will cause a division by zero error (#DIV/0!)
            Cell errorCell = sheet.Cells["A1"];
            errorCell.Formula = "=10/0";

            // Force calculation to generate the error in the cell
            workbook.CalculateFormula();

            // Verify that the cell indeed contains an error
            if (errorCell.Type == CellValueType.IsError)
            {
                Console.WriteLine("Source cell contains an error as expected.");
            }
            else
            {
                Console.WriteLine("Unexpected: source cell does not contain an error.");
            }

            // Prepare PDF save options (no special error‑ignoring flag is required)
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Export the workbook to PDF
            workbook.Save("ExportedWithIgnoreErrors.pdf", saveOptions);
            Console.WriteLine("Export completed successfully without throwing exceptions.");
        }
        catch (Exception ex)
        {
            // Any exception during processing will be reported here
            Console.WriteLine("Operation failed with exception: " + ex.Message);
        }
    }
}
