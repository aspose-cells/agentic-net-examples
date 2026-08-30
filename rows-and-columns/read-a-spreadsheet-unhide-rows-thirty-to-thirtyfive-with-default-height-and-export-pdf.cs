// Title: How to unhide rows 30‑35 with default height and save the worksheet as PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to unhide rows 30 through 35 (default height) and then export the workbook to a PDF file. | Create a reusable C# method that takes a worksheet, start row, row count, and height, unhides those rows with Aspose.Cells, and returns a PDF stream. | Adapt the Aspose.Cells example to accept a user‑provided Excel file path, unhide a configurable row range, and save the result as a PDF.
// Common Searches: Aspose.Cells C# unhide rows 30-35 before converting to PDF | set default row height while unhiding rows using Aspose.Cells .NET | convert Excel to PDF after making hidden rows visible with Aspose.Cells | C# code to unhide specific rows in a worksheet and export to PDF using Aspose.Cells
// Tags: Aspose.Cells row visibility C# | Aspose.Cells PDF export .NET | default row height Aspose.Cells | worksheet row range manipulation Aspose.Cells | unhide rows operation Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an existing Excel file with Aspose.Cells, unhides rows 30‑35 using the default height (-1), and saves the worksheet as a PDF document.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Unhide rows 30 to 35 (zero‑based index), total 6 rows, default height (-1)
        workbook.Worksheets[0].Cells.UnhideRows(30, 6, -1);

        // Export the workbook to PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
