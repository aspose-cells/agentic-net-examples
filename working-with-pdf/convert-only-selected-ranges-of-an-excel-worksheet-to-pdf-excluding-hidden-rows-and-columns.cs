// Title: Convert a Selected Excel Range to PDF While Ignoring Hidden Rows and Columns – Aspose.Cells for .NET
// Description: This example demonstrates how to load an Excel workbook with Aspose.Cells, hide specific rows and columns, define a print area (e.g., A1:D20), and save only that visible range to a PDF using PdfSaveOptions. Hidden rows and columns are automatically excluded from the output.
// Keywords: Aspose.Cells export range to PDF | ignore hidden rows PDF Aspose.Cells | ignore hidden columns PDF Aspose.Cells | set print area PDF conversion .NET | PdfSaveOptions selected range | C# convert Excel to PDF | Aspose.Cells hide rows columns PDF
// Common Searches: Aspose.Cells export only A1:D20 to PDF | How to exclude hidden rows when saving Excel as PDF with Aspose | Set print area for PDF output using Aspose.Cells .NET | Convert Excel worksheet to PDF without hidden columns | C# Aspose.Cells PDF conversion selected range
// Developer Intent: Generate a PDF that contains only a specified cell range from an Excel worksheet, automatically omitting any rows or columns that are hidden.
// Use Cases: Create a clean report PDF that shows only the visible portion of a data table. | Produce printable invoices from a predefined area of a worksheet while ignoring hidden formatting rows. | Automate batch processing of multiple sheets, each with its own print area, to generate PDFs that contain only visible data.
// AI Prompts: Write C# code with Aspose.Cells to export the A1:D20 range of a worksheet to PDF, ensuring hidden rows and columns are excluded. | Show how to set a print area and use PdfSaveOptions to save only visible cells to a PDF file in .NET. | Explain how Aspose.Cells determines which hidden rows or columns are omitted when converting a selected range to PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example demonstrates how to load an Excel workbook with Aspose.Cells, hide specific rows and columns, define a print area (e.g., A1:D20), and save only that visible range to a PDF using PdfSaveOptions. Hidden rows and columns are automatically excluded from the output.
class ConvertSelectedRangeToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output_selected_range.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Example: hide some rows and columns.
            sheet.Cells.HideRow(2);      // Hide row 3 (zero‑based index)
            sheet.Cells.HideColumn(1);  // Hide column B

            // Define the range to export (e.g., A1:D20) and set it as the print area.
            string exportRange = "A1:D20";
            sheet.PageSetup.PrintArea = exportRange;

            // Create PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Note: In older Aspose.Cells versions the IgnoreHiddenRows/Columns
            // properties are not available. Hidden rows/columns are excluded
            // from the PDF when a print area is defined, so no additional settings are required.

            // Save the selected range to PDF.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
