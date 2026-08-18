// Title: C# – Hide Columns D‑G in an Excel Worksheet and Export to PDF using Aspose.Cells
// Description: Load an Excel file with Aspose.Cells, hide columns D through G (zero‑based indices 3‑6) on the first worksheet, and save the result directly as a PDF. The example demonstrates the Cells.HideColumns method and PDF export in .NET.
// Keywords: Aspose.Cells hide columns C# | hide columns D to G | Excel to PDF conversion .NET | Cells.HideColumns example | export worksheet as PDF Aspose | C# Excel column visibility | Aspose.Cells PDF output
// Common Searches: how to hide a range of columns in Aspose.Cells C# | Aspose.Cells hide columns D‑G and save as PDF | C# convert Excel to PDF after hiding columns | Aspose.Cells hide multiple columns before PDF export
// Developer Intent: Hide a specific column range in an Excel sheet and generate a PDF of the modified worksheet.
// Use Cases: Produce printable reports that exclude confidential columns. | Create clean PDF versions of spreadsheets for client distribution. | Automate batch processing to hide designated columns across many workbooks before PDF conversion.
// AI Prompts: Show C# code that hides columns D‑G in an Aspose.Cells workbook and exports the sheet to PDF. | Give an Aspose.Cells example for hiding a dynamic column range based on letters and then saving as PDF. | Explain how to hide columns while preserving their widths during Excel‑to‑PDF conversion with Aspose.Cells.

using System;
using Aspose.Cells;

// Load an Excel file with Aspose.Cells, hide columns D through G (zero‑based indices 3‑6) on the first worksheet, and save the result directly as a PDF. The example demonstrates the Cells.HideColumns method and PDF export in .NET.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Hide columns D (index 3) through G (index 6)
        // Total columns to hide = 4 (D, E, F, G)
        int startColumn = 3;      // Column D (zero‑based)
        int columnCount = 4;      // D, E, F, G
        cells.HideColumns(startColumn, columnCount);

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
