// Title: How to unhide rows 40‑45, set a custom row height, and export an Excel worksheet to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells unhide rows 40 to 45 and set row height before PDF export | set custom row height for specific rows in Aspose.Cells and save as PDF | make hidden rows visible in an Excel workbook with Aspose.Cells .NET | Aspose.Cells export to PDF after adjusting row visibility and height | unhide multiple rows and define height in Aspose.Cells C# example
// Tags: unhide rows Aspose.Cells C# | custom row height Aspose.Cells | export Excel to PDF Aspose.Cells | modify row visibility Aspose.Cells | row height adjustment PDF conversion Aspose.Cells

using Aspose.Cells;
using System;

// The sample loads an Excel file, accesses the first worksheet, iterates rows 40‑45 (zero‑based indices 39‑44) to make each row visible and set its height to 25 points, and then saves the workbook as a PDF document.
class Program
{
    static void Main()
    {
        // Load the Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Unhide rows 40 to 45 (zero‑based index 39‑44) and set a custom height
        for (int i = 39; i <= 44; i++)
        {
            Row row = sheet.Cells.Rows[i];
            row.IsHidden = false;          // make the row visible
            row.Height = 25;               // set custom height (points)
        }

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
