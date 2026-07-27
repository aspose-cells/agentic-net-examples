using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put a value that starts with a single quote (apostrophe)
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("'12345"); // The leading apostrophe is part of the cell's text

        // Create a style (default style) and set its QuotePrefix to true
        // This demonstrates that the style could change the QuotePrefix if applied
        Style style = workbook.CreateStyle();
        style.QuotePrefix = true;

        // Create a StyleFlag and set QuotePrefix to false
        // This tells SetStyle not to modify the QuotePrefix property of the cell
        StyleFlag flag = new StyleFlag();
        flag.QuotePrefix = false;

        // Apply the style to the cell using the flag
        // Because QuotePrefix flag is false, the existing apostrophe prefix is preserved
        cell.SetStyle(style, flag);

        // Save the workbook
        workbook.Save("QuotePrefixPreserve.xlsx");
    }
}