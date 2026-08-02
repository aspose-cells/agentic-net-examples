// Title: Freeze Top Row After AutoFitRows to Preserve Height – Aspose.Cells C# Example
// Description: A concise C# sample that builds a workbook, inserts wrapped text, auto‑fits every row, then freezes the top row so its height remains unchanged while scrolling, finally saving the result as an XLSX file.
// Keywords: Aspose.Cells | C# | FreezePanes | AutoFitRows | preserve row height | scrolling | wrap text | Excel export | worksheet freeze top row | row height lock
// Common Searches: Aspose.Cells freeze top row after autofitrows | keep row height fixed while scrolling Excel C# | how to apply FreezePanes after AutoFitRows in .NET | lock header row height with Aspose.Cells | C# example for AutoFitRows then FreezePanes
// Developer Intent: Apply FreezePanes after AutoFitRows so the first row’s height stays constant during scrolling.
// Use Cases: Financial reports where the header row must stay visible and retain its auto‑fitted height. | Invoice sheets with wrapped description cells that are auto‑sized, while the top row remains frozen for quick reference. | Large data exports where readability is improved by freezing the first row after adjusting row heights.
// AI Prompts: Generate C# code using Aspose.Cells that wraps text, auto‑fits rows, then freezes the top row to keep its height static while scrolling. | Show an Aspose.Cells example that calls Worksheet.AutoFitRows before Worksheet.FreezePanes and saves the workbook to a custom path. | Explain why the FreezePanes method should be invoked after AutoFitRows to maintain row height in an Excel file created with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace FreezePanesAfterAutoFitRowsDemo
{
    // A concise C# sample that builds a workbook, inserts wrapped text, auto‑fits every row, then freezes the top row so its height remains unchanged while scrolling, finally saving the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data that will affect row heights
            sheet.Cells["A1"].PutValue("Short text");
            sheet.Cells["A2"].PutValue("This is a longer piece of text that should cause the row height to increase when wrapped.");
            sheet.Cells["A3"].PutValue("Another line with\nmultiple line breaks\nto demonstrate auto‑fit.");

            // Enable text wrapping for the cells to allow row height changes
            Style wrapStyle = sheet.Cells["A2"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            sheet.Cells["A2"].SetStyle(wrapStyle);
            sheet.Cells["A3"].SetStyle(wrapStyle);

            // Auto‑fit all rows so that their heights match the content
            sheet.AutoFitRows();

            // Freeze the first row (row index 1) to lock its height during scrolling
            // Parameters: row index, column index, number of frozen rows, number of frozen columns
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook to the desktop (adjust the path as needed)
            string outputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "FreezePanesAfterAutoFitRows.xlsx");

            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
