using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access a cell and set a value that starts with a single quote (apostrophe)
            Cell cell = worksheet.Cells["B2"];
            cell.PutValue("'12345"); // The leading apostrophe forces Excel to treat the value as text

            // Create a style and set QuotePrefix to false (remove the leading apostrophe)
            Style style = workbook.CreateStyle();
            style.QuotePrefix = false;

            // Create a StyleFlag and enable the QuotePrefix flag so the change is applied
            StyleFlag flag = new StyleFlag();
            flag.QuotePrefix = true;

            // Apply the style to the cell using the flag (lifecycle: set style)
            cell.SetStyle(style, flag);

            // Save the workbook (lifecycle: save)
            workbook.Save("QuotePrefixRemoved.xlsx", SaveFormat.Xlsx);

            // Optional: Verify the change by reading the property back
            Workbook loaded = new Workbook("QuotePrefixRemoved.xlsx");
            Cell loadedCell = loaded.Worksheets[0].Cells["B2"];
            Console.WriteLine("QuotePrefix after removal: " + loadedCell.GetStyle().QuotePrefix);
            Console.WriteLine("Cell value: " + loadedCell.StringValue);
        }
    }
}