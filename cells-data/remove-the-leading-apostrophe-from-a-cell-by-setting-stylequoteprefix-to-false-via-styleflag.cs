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
            Cells cells = sheet.Cells;

            // Put a value that starts with a single quote (apostrophe)
            Cell cell = cells["B2"];
            cell.PutValue("'12345"); // Excel treats this as text with QuotePrefix = true

            // Create a style and set QuotePrefix to false (remove the leading apostrophe)
            Style style = workbook.CreateStyle();
            style.QuotePrefix = false;

            // Create a StyleFlag and enable the QuotePrefix flag so the change is applied
            StyleFlag flag = new StyleFlag();
            flag.QuotePrefix = true;

            // Apply the style to the cell using the flag
            cell.SetStyle(style, flag);

            // Verify the change (should be false)
            Console.WriteLine("QuotePrefix after removal: " + cell.GetStyle().QuotePrefix);
            Console.WriteLine("Cell value after removal: " + cell.StringValue);

            // Save the workbook
            workbook.Save("QuotePrefixRemoved.xlsx");
        }
    }
}