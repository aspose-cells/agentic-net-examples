// Title: Remove Empty Columns and Convert Excel to PDF with Aspose.Cells (C#)
// Description: Loads an Excel workbook, deletes every column that contains no data using Worksheet.Cells.DeleteBlankColumns(), and saves the cleaned worksheet as a PDF via SaveFormat.Pdf.
// Keywords: Aspose.Cells DeleteBlankColumns | C# remove empty columns Excel | Aspose.Cells export to PDF | Excel column cleanup C# | Aspose.Cells PDF conversion | delete blank columns Aspose | Aspose.Cells .NET PDF
// Common Searches: Aspose.Cells delete empty columns C# | How to remove blank columns before PDF export using Aspose.Cells | Convert trimmed Excel sheet to PDF with Aspose.Cells .NET | Worksheet.Cells.DeleteBlankColumns example | C# code to clean Excel columns and save as PDF
// Developer Intent: The developer wants to strip all blank columns from an Excel sheet and then generate a PDF of the resulting worksheet.
// Use Cases: Prepare client‑ready reports by eliminating unused columns before creating a PDF. | Automate data pipelines so only populated columns appear in the final PDF document. | Archive Excel workbooks in a compact PDF format after removing empty columns.
// AI Prompts: Generate C# code that uses Aspose.Cells to delete blank columns from a worksheet and export the result to PDF, handling multiple worksheets if needed. | Explain how DeleteBlankColumns works and how to retain column formatting when converting to PDF with Aspose.Cells. | Create a reusable method that accepts input and output paths, removes empty columns, saves as PDF, and returns a success status using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an Excel workbook, deletes every column that contains no data using Worksheet.Cells.DeleteBlankColumns(), and saves the cleaned worksheet as a PDF via SaveFormat.Pdf.
    class DeleteBlankColumnsAndConvertToPdf
    {
        static void Main()
        {
            // Load the source Excel file
            Workbook workbook = new Workbook("input.xlsx");

            // Work with the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Remove all columns that contain no data
            worksheet.Cells.DeleteBlankColumns();

            // Save the resulting worksheet as a PDF document
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
