// Title: C# – Hide Column C in an Excel Worksheet and Export to PDF using Aspose.Cells
// Description: Load an existing workbook, hide column C (index 2) on the first worksheet with Aspose.Cells, then save the result directly as a PDF file.
// Keywords: Aspose.Cells hide column C | C# export Excel to PDF | Aspose.Cells column visibility | convert Excel to PDF .NET | hide column before PDF conversion
// Common Searches: Aspose.Cells hide column C example | C# hide Excel column and save as PDF | how to hide a column with Aspose.Cells | export hidden‑column Excel to PDF .NET | Aspose.Cells PDF conversion after hiding column
// Developer Intent: Hide column C in the first worksheet and generate a PDF from the modified workbook.
// Use Cases: Produce client‑ready PDFs that exclude internal data stored in a specific column. | Automate batch reporting where confidential columns must be hidden before distribution. | Create printable PDFs from Excel templates while selectively omitting unnecessary columns.
// AI Prompts: Generate C# code that hides column D in the second worksheet and saves the workbook as a PDF with Aspose.Cells. | Show how to hide multiple columns (B, C, and E) and then export the worksheet to PDF using Aspose.Cells for .NET. | Provide an example that hides a column, changes its width, and converts the sheet to PDF while preserving all other formatting.

using System;
using Aspose.Cells;

// Load an existing workbook, hide column C (index 2) on the first worksheet with Aspose.Cells, then save the result directly as a PDF file.
class Program
{
    static void Main()
    {
        // Path to the existing Excel file
        string inputPath = "input.xlsx";

        // Desired PDF output path
        string outputPath = "output.pdf";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(inputPath);

        // Hide column C (zero‑based index 2) in the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells.HideColumn(2);

        // Save the modified workbook as a PDF file
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}
