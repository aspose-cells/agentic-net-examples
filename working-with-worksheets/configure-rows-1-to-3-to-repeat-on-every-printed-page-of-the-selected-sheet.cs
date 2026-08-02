// Title: Set rows 1‑3 as repeatable print titles on every page with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills sample data, and uses Worksheet.PageSetup.PrintTitleRows = "$1:$3" so that rows 1‑3 repeat on each printed page. Saves the file as RepeatRowsDemo.xlsx.
// Keywords: Aspose.Cells | C# | PrintTitleRows | repeat rows on printed page | Excel header rows | page setup print titles | worksheet print titles | Excel to PDF printing | Aspose.Cells API
// Common Searches: Aspose.Cells repeat rows on each printed page | How to set PrintTitleRows in C# | Configure page setup to repeat header rows in Excel using Aspose | Aspose.Cells print titles rows example | C# code for repeating rows when printing Excel
// Developer Intent: Apply the PrintTitleRows property so rows 1‑3 are printed as titles on every page of the worksheet.
// Use Cases: Generate multi‑page reports where the first three rows contain titles or column headings that must appear on each printed sheet. | Create printable Excel invoices with consistent header rows across all pages. | Automate workbook creation for PDF conversion while preserving repeated header rows.
// AI Prompts: Show C# code to repeat rows 1‑5 and columns A‑C on each printed page with Aspose.Cells. | Explain how to validate the PrintTitleRows setting after saving a workbook. | Provide a full example that configures both PrintTitleRows and PrintTitleColumns for a worksheet.

using System;
using Aspose.Cells;

// Creates a workbook, fills sample data, and uses Worksheet.PageSetup.PrintTitleRows = "$1:$3" so that rows 1‑3 repeat on each printed page. Saves the file as RepeatRowsDemo.xlsx.
class RepeatRowsPrintTitle
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data (optional, just to illustrate the effect)
        for (int i = 0; i < 20; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Configure rows 1 to 3 to repeat on every printed page
        worksheet.PageSetup.PrintTitleRows = "$1:$3";

        // Save the workbook
        workbook.Save("RepeatRowsDemo.xlsx");
    }
}
