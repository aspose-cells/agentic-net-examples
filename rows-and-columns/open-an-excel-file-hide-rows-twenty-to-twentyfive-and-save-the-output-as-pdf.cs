// Title: Hide rows 20‑25 in an Excel file with Aspose.Cells for .NET and export the workbook to PDF
// AI Prompts: Generate C# code that opens an existing .xlsx file, hides rows 20 through 25 on the first worksheet using Aspose.Cells, and saves the result as a PDF document. | Write a C# snippet that uses Aspose.Cells to conceal a specific range of rows in an Excel workbook and then export the modified workbook to PDF.
// Common Searches: how to hide a range of rows in an Excel workbook using Aspose.Cells C# | Aspose.Cells hide rows then save as PDF example | C# code to conceal rows 20 to 25 before PDF conversion | export Excel worksheet to PDF after hiding rows with Aspose.Cells
// Tags: Aspose.Cells hide rows | Aspose.Cells PDF export | C# hide Excel rows | Aspose.Cells row visibility | Excel to PDF conversion Aspose.Cells

using System;
using Aspose.Cells;

// Loads input.xlsx, hides rows 20‑25 in the first worksheet via HideRows, and saves the workbook as output.pdf in PDF format.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Hide rows 20 to 25 (zero‑based index: start at 19, hide 6 rows)
        workbook.Worksheets[0].Cells.HideRows(19, 6);

        // Save the modified workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
