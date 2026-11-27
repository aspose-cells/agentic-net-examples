// Title: How to unhide columns N‑P, set a 70‑point width, and export a worksheet to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unhide specific column range before PDF export | set column width in points with Aspose.Cells .NET prior to saving as PDF | unhide columns N to P in an Excel worksheet using Aspose.Cells and generate PDF | adjust column visibility and width in Aspose.Cells workbook then export to PDF | C# code to reveal multiple hidden columns and define their width in points with Aspose.Cells
// Tags: unhide column range Aspose.Cells .NET | set column width points Aspose.Cells | export worksheet to PDF Aspose.Cells | column visibility manipulation Aspose.Cells | column width configuration PDF export

using Aspose.Cells;
using System;

// The example loads an Excel workbook, accesses the first worksheet, unhides columns N through P (indices 13‑15), sets each column's width to 70 points, and saves the workbook as a PDF file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Unhide columns N (index 13) through P (index 15) and set their width to 70 points
        for (int col = 13; col <= 15; col++)
        {
            Column column = sheet.Cells.Columns[col];
            column.IsHidden = false;
            column.Width = 70; // Approximate width in points
        }

        // Save the worksheet as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
