// Title: C# example: hide column C in an Excel worksheet and save the workbook as a PDF using Aspose.Cells
// Common Searches: Aspose.Cells C# hide column C in Excel before PDF conversion | How to hide a specific column in a worksheet and export to PDF with Aspose.Cells .NET | C# code to hide column index 2 in Excel and generate PDF using Aspose.Cells | Export Excel to PDF while keeping column C hidden using Aspose.Cells for .NET
// Tags: hide column Aspose.Cells C# | Excel column visibility PDF export Aspose.Cells | Aspose.Cells hide column index 2 | convert workbook to PDF with hidden columns Aspose.Cells | C# Aspose.Cells hide column before PDF save

using System;
using Aspose.Cells;

// Loads 'input.xlsx', hides column C (zero‑based index 2) in the first worksheet, and saves the modified workbook as 'output.pdf' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Hide column C (zero‑based index 2) in the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.HideColumn(2);

        // Save the modified workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
