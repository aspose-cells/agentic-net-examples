// Title: Hide rows 21‑25 in an Excel worksheet and export the visible view to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells to conceal rows 21‑25 in the first worksheet and then generate a PDF that reflects the hidden rows. | Create a script that opens an .xlsx file, marks rows 21‑25 as hidden, and saves the current worksheet view as a PDF using Aspose.Cells.
// Common Searches: Aspose.Cells C# conceal rows 21‑25 before PDF conversion | How to export an Excel sheet to PDF with specific rows hidden using Aspose.Cells | C# code to hide a range of rows in a workbook and save as PDF with Aspose.Cells | Save Excel view with hidden rows as PDF using Aspose.Cells .NET
// Tags: Aspose.Cells hide rows C# | Aspose.Cells export PDF hidden rows | C# hide Excel rows Aspose.Cells | Aspose.Cells PDF conversion with hidden rows

using System;
using Aspose.Cells;

// Loads input.xlsx, hides rows 21‑25 on the first worksheet, and saves the visible view as output.pdf using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Hide rows 21 to 25 (zero‑based indices 20‑24)
        for (int rowIndex = 20; rowIndex <= 24; rowIndex++)
        {
            sheet.Cells.Rows[rowIndex].IsHidden = true;
        }

        // Save the current view of the workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
