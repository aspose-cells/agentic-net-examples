// Title: Hide columns C and D in an Excel workbook and export to PDF with Aspose.Cells for .NET
// Common Searches: C# hide specific columns in Excel with Aspose.Cells and export to PDF | Aspose.Cells hide column C and D before saving workbook as PDF | Preserve hidden column settings when converting Excel to PDF using Aspose.Cells .NET | How to hide columns in a worksheet and generate PDF with Aspose.Cells C# | Export Excel to PDF while keeping columns hidden using Aspose.Cells library
// Tags: hide columns Aspose.Cells C# | export workbook to PDF Aspose.Cells | preserve hidden column state PDF conversion | worksheet column visibility control Aspose.Cells | Excel to PDF conversion with hidden columns

using Aspose.Cells;
using System;

// The example loads an existing Excel file, hides columns C and D on the first worksheet using Aspose.Cells, and then saves the workbook as a PDF. The hidden columns remain omitted in the generated PDF, preserving the worksheet's column visibility settings.
class Program
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (you can change the index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Hide column C (zero‑based index 2) and column D (index 3)
        sheet.Cells.HideColumn(2);
        sheet.Cells.HideColumn(3);

        // Export the workbook while preserving hidden columns.
        // When exporting to PDF, hidden columns are not rendered.
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
