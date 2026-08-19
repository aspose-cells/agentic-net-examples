// Title: Add Apostrophe Prefix to Numeric Cells Using QuotePrefix in Aspose.Cells (C#)
// Description: Creates a workbook, inserts numeric values, scans all used cells, detects numeric types, applies a Style with QuotePrefix enabled via StyleFlag, and saves the file so the numbers are stored as text with a leading apostrophe.
// Keywords: Aspose.Cells QuotePrefix | C# numeric to text conversion | Excel apostrophe prefix programmatically | force text format for numbers | preserve leading zeros Aspose.Cells | prevent scientific notation Excel C#
// Common Searches: Aspose.Cells add apostrophe to numeric cells | QuotePrefix flag C# Aspose.Cells example | convert numbers to text in Excel using Aspose | how to force text format for numbers Aspose.Cells | apply style flag QuotePrefix programmatically
// Developer Intent: Enable QuotePrefix on numeric cells so they are saved as text with an apostrophe prefix.
// Use Cases: Keep leading zeros in ID columns when exporting data. | Avoid scientific notation for large numeric identifiers. | Store account numbers, product codes, or ZIP codes as exact text.
// AI Prompts: Generate C# code that iterates a worksheet and adds an apostrophe prefix to every numeric cell using Aspose.Cells QuotePrefix. | Create a reusable method that applies a QuotePrefix style only to cells of type IsNumeric, leaving other cells unchanged. | Explain the interaction between Style, StyleFlag, and the QuotePrefix property for forcing text representation of numbers in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    // Creates a workbook, inserts numeric values, scans all used cells, detects numeric types, applies a Style with QuotePrefix enabled via StyleFlag, and saves the file so the numbers are stored as text with a leading apostrophe.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data: numeric values that should be treated as text
            cells["A1"].PutValue(12345);          // integer
            cells["A2"].PutValue(9876.54);        // double
            cells["B1"].PutValue("NormalText");   // non‑numeric (should remain unchanged)

            // Iterate through all used cells
            foreach (Cell cell in cells)
            {
                // Check if the cell currently holds a numeric value
                if (cell.Type == CellValueType.IsNumeric)
                {
                    // Create a style with QuotePrefix enabled
                    Style style = workbook.CreateStyle();
                    style.QuotePrefix = true;

                    // Enable the QuotePrefix flag so it is applied
                    StyleFlag flag = new StyleFlag();
                    flag.QuotePrefix = true;

                    // Apply the style to the cell
                    cell.SetStyle(style, flag);
                }
            }

            // Save the workbook – the numeric cells will now be stored as text (prefixed with an apostrophe)
            workbook.Save("NumericCellsWithApostrophe.xlsx");
        }
    }
}
