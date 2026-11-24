// Title: Unhide columns H‑J, set each to a 50‑point width, and export the workbook to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# unhide columns H J and set column width | Set column width in points with Aspose.Cells before PDF export | Export Excel to PDF after adjusting column visibility using Aspose.Cells | C# code to change column visibility and width then save as PDF with Aspose.Cells | How to convert Excel to PDF with specific column widths in Aspose.Cells .NET
// Tags: unhide specific columns Aspose.Cells | set column width in points Aspose.Cells | convert Excel to PDF Aspose.Cells | adjust column visibility before PDF export Aspose.Cells | column width pixel conversion Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an existing Excel workbook, unhides columns H through J on the first worksheet, sets each column's width to approximately 50 points (using pixel conversion), and then saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx"); // replace with your source file path

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Unhide columns H (index 7) through J (index 9) and set their width to 50 points
        // Aspose.Cells column width is measured in characters; to approximate 50 points,
        // we can use SetColumnWidthPixel (1 point ≈ 1.333 pixels at 96 DPI).
        // 50 points ≈ 67 pixels.
        for (int col = 7; col <= 9; col++)
        {
            sheet.Cells.Columns[col].IsHidden = false;               // unhide column
            sheet.Cells.SetColumnWidthPixel(col, 67);                // set width ~50 points
        }

        // Save the workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf); // replace with desired output path
    }
}
