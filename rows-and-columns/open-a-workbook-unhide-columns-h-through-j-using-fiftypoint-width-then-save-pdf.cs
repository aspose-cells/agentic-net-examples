// Title: C# – Unhide Columns H‑J (50 pt) and Export Workbook to PDF with Aspose.Cells
// Description: Load an Excel file, reveal columns H, I and J, set each column width to 50 points using the UnhideColumns method, and save the worksheet directly as a PDF. Demonstrates Aspose.Cells for .NET column formatting and PDF conversion in a single script.
// Keywords: Aspose.Cells C# unhide columns | UnhideColumns method | set column width points | export Excel to PDF .NET | column formatting Aspose.Cells | C# Excel PDF conversion | Aspose.Cells column visibility
// Common Searches: how to unhide specific columns with Aspose.Cells C# | set column width in points using Aspose.Cells | convert Excel workbook to PDF after changing column visibility | Aspose.Cells example unhide columns H to J | C# code to export hidden columns to PDF
// Developer Intent: Reveal columns H‑J, assign a 50‑point width to each, and generate a PDF from the workbook.
// Use Cases: Create printable reports where hidden data columns must appear with uniform width. | Automate PDF generation for financial statements that require certain columns to be visible and sized consistently. | Prepare marketing dashboards where summary columns need to be displayed at a fixed width before export.
// AI Prompts: Write C# code that unhides columns H through J, sets each column width to 50 points, and saves the workbook as a PDF using Aspose.Cells. | Explain the parameters of the UnhideColumns method and how to use SaveFormat.Pdf for Excel‑to‑PDF conversion. | Show how to adjust column visibility and width in Aspose.Cells before converting a worksheet to PDF.

using System;
using Aspose.Cells;

// Load an Excel file, reveal columns H, I and J, set each column width to 50 points using the UnhideColumns method, and save the worksheet directly as a PDF. Demonstrates Aspose.Cells for .NET column formatting and PDF conversion in a single script.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Unhide columns H (index 7) through J (index 9) and set each column width to 50 points
        // Parameters: start column index, number of columns, width in points
        worksheet.Cells.UnhideColumns(7, 3, 50);

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
