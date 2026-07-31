// Title: AutoFit rows before FreezePanes in Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, insert wrapped and multi‑line text, call AutoFitRows to set correct row heights, then freeze the top row and first column with FreezePanes, and save the result as an XLSX file. The row heights stay consistent while scrolling.
// Keywords: Aspose.Cells | C# | .NET | AutoFitRows | FreezePanes | row height | text wrapping | Excel export | worksheet freeze | auto fit rows before freeze
// Common Searches: Aspose.Cells AutoFitRows before FreezePanes | C# auto fit rows then freeze panes | preserve wrapped text row height Aspose.Cells | how to freeze top row after auto fitting rows .NET | auto fit rows and freeze panes example
// Developer Intent: Adjust all row heights with AutoFitRows before invoking FreezePanes so wrapped or multi‑line cells keep their expanded size while the pane is frozen.
// Use Cases: Generate a report where header rows stay visible and wrapped text cells retain their height after freezing the first row and column. | Export a spreadsheet containing multi‑line comments that must keep proper row height while users navigate frozen panes. | Create a reusable template that automatically fits rows with wrapped content and then freezes panes to preserve layout stability in Excel.
// AI Prompts: Write C# code using Aspose.Cells to enable text wrapping, auto‑fit all rows, and freeze the first row and column. | Explain why AutoFitRows should be called before FreezePanes to maintain row height in an Aspose.Cells workbook. | Provide a step‑by‑step guide for auto‑fitting rows with wrapped text and then freezing panes in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsAutoFitAndFreeze
{
    // Shows how to create a workbook, insert wrapped and multi‑line text, call AutoFitRows to set correct row heights, then freeze the top row and first column with FreezePanes, and save the result as an XLSX file. The row heights stay consistent while scrolling.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data with wrapped text to demonstrate row height changes
            sheet.Cells["A1"].PutValue("This is a long text that will require the row to expand when AutoFitRows is applied.");
            sheet.Cells["A2"].PutValue("Short text");
            sheet.Cells["A3"].PutValue("Another long text\nwith line breaks\nto test multi‑line row height.");

            // Enable text wrapping for the cells that need it
            Style wrapStyle = sheet.Cells["A1"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            sheet.Cells["A1"].SetStyle(wrapStyle);

            wrapStyle = sheet.Cells["A3"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            sheet.Cells["A3"].SetStyle(wrapStyle);

            // Auto‑fit all rows in the worksheet before freezing panes
            sheet.AutoFitRows();

            // Freeze the top row (row index 1) and first column (column index 1)
            // Parameters: row index, column index, number of frozen rows, number of frozen columns
            sheet.FreezePanes(1, 1, 1, 1);

            // Save the workbook to the desktop
            string outputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "AutoFitRowsAndFreezePanes.xlsx");

            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
