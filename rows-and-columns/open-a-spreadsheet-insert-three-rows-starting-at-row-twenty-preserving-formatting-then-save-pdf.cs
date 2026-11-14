// Title: Insert three rows at Excel row 20 while preserving formatting and export the workbook to PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# insert rows at a specific row and keep original formatting | How to add multiple rows at row 20 in an Excel file with Aspose.Cells and retain styles | Save modified Excel workbook as PDF after inserting rows using Aspose.Cells .NET | Preserve row formatting when inserting rows with Aspose.Cells in C# | Insert rows starting at row 20 and export to PDF with Aspose.Cells for .NET
// Tags: insert rows preserve formatting Aspose.Cells | Aspose.Cells InsertRows method C# | export workbook to PDF Aspose.Cells | add rows at specific index Excel .NET | preserve Excel row style Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an existing Excel file (input.xlsx), inserts three rows beginning at row 20 while preserving the original formatting, and then saves the updated workbook as a PDF (output.pdf) using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Insert three rows starting at Excel row 20 (zero‑based index 19)
        // The third parameter 'true' preserves the formatting of the inserted rows
        sheet.Cells.InsertRows(19, 3, true);

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
