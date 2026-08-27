// Title: Hide columns D‑G in an Excel worksheet and export the sheet as PDF using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells in C# to hide a contiguous block of columns (e.g., columns 4‑7) and then save the first worksheet as a PDF. | Programmatically conceal a specific column range in an Excel file and generate a PDF output with the Aspose.Cells API.
// Common Searches: C# Aspose.Cells hide column range before PDF conversion | How to hide multiple Excel columns and export to PDF using Aspose.Cells .NET | Aspose.Cells hide columns D to G and create PDF example | Hide specific columns in Excel workbook then convert to PDF with Aspose.Cells
// Tags: Aspose.Cells column visibility control | Aspose.Cells export worksheet to PDF | Cells.HideColumns method usage | PDF generation after hiding columns Aspose.Cells | Excel column concealment .NET

using System;
using Aspose.Cells;

// The sample loads 'input.xlsx' into a Workbook, hides columns D through G on the first worksheet using Cells.HideColumns, and then saves the modified workbook as 'output.pdf' in PDF format.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Hide columns D (index 3) through G (index 6) – total of 4 columns
        int startColumn = 3;      // Column D
        int totalColumns = 4;     // D, E, F, G
        cells.HideColumns(startColumn, totalColumns);

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
