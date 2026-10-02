// Title: How to unhide rows 40‑45, set a custom height, and save the worksheet as PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, makes rows 40‑45 visible, assigns each a height of 25 points, and saves the workbook as a PDF. | Generate a snippet that iterates over rows 40‑45 in a worksheet, clears the hidden flag, applies a custom row height, and converts the workbook to PDF using Aspose.Cells.
// Common Searches: Aspose.Cells C# unhide specific rows before PDF export | Set row height for a range of rows in Excel and convert to PDF using Aspose.Cells | How to make hidden rows visible in a .NET workbook and save as PDF | C# code to adjust row visibility and height then export Excel to PDF with Aspose.Cells | Unhide rows 40 to 45 and set custom height in Aspose.Cells before PDF conversion
// Tags: unhide rows Aspose.Cells C# | apply row height Aspose.Cells | export workbook to PDF Aspose.Cells | adjust row visibility C# | range row formatting Aspose.Cells

using Aspose.Cells;
using System;

// The example loads 'input.xlsx', accesses the first worksheet, iterates over rows 40‑45 (zero‑based indices 39‑44), clears the hidden flag, sets each row's height to 25 points, and saves the workbook as 'output.pdf' in PDF format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Unhide rows 40 to 45 (1‑based) and set a custom height
        for (int i = 39; i <= 44; i++) // zero‑based row index
        {
            Row row = sheet.Cells.Rows[i];
            row.IsHidden = false;      // make the row visible
            row.Height = 25;           // set custom height (points)
        }

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
