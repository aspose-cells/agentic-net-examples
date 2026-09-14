// Title: How to auto‑fit a single worksheet row with wrapped text using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that writes a long string to cell A1, enables text wrapping, and calls Worksheet.AutoFitRow(0) with Aspose.Cells. | Create a .NET snippet that adjusts the height of a specific row after applying a wrap‑text style to its cells using Aspose.Cells. | Provide an example that demonstrates auto‑fitting multiple rows individually after setting different cell styles in Aspose.Cells C#.
// Common Searches: Aspose.Cells C# auto fit row height after enabling text wrap | Worksheet.AutoFitRow method usage example for a single row | How to adjust row height based on long text in Aspose.Cells .NET | C# code to auto‑size row 1 in an Excel file with Aspose.Cells
// Tags: Worksheet.AutoFitRow C# example | auto-fit row height Aspose.Cells .NET | wrap text cell style Aspose.Cells | adjust row height based on content Aspose.Cells | Excel row auto-fit using Aspose.Cells

using System;
using Aspose.Cells;

namespace AutoFitRowExample
{
    // // Demonstrates creating a workbook, inserting a long wrapped text into cell A1, enabling text wrapping, auto‑fitting the first row with Worksheet.AutoFitRow, and saving the file as AutoFitRowResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add long text to a cell to require row height adjustment
            worksheet.Cells["A1"].PutValue("This is a very long piece of text that should cause the row to expand when auto‑fitted.");
            // Enable text wrapping so the content occupies multiple lines
            Style style = worksheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(style);

            // Auto‑fit the first row (row index 0) based on its content
            worksheet.AutoFitRow(0);

            // Save the workbook (save rule)
            workbook.Save("AutoFitRowResult.xlsx");
        }
    }
}
