using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a value that starts with a single quote (apostrophe)
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("'SampleText");

            // Create a style (no changes needed, just a placeholder)
            Style style = workbook.CreateStyle();

            // Create a StyleFlag and set QuotePrefix to false.
            // This tells SetStyle not to modify the QuotePrefix property of the cell,
            // thereby preserving the existing apostrophe prefix.
            StyleFlag flag = new StyleFlag
            {
                QuotePrefix = false
            };

            // Apply the style to the cell using the flag
            cell.SetStyle(style, flag);

            // Save the workbook to verify the result
            workbook.Save("QuotePrefixPreserved.xlsx");
        }
    }
}