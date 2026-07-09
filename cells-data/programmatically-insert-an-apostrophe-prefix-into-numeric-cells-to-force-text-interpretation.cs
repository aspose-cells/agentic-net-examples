using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some numeric values
            cells["A1"].PutValue(12345);
            cells["A2"].PutValue(67890);
            cells["B1"].PutValue("AlreadyText"); // non‑numeric, will stay unchanged

            // Apply QuotePrefix style to numeric cells to force text interpretation
            // (uses Style.QuotePrefix and StyleFlag.QuotePrefix rules)
            ApplyQuotePrefix(cells["A1"]);
            ApplyQuotePrefix(cells["A2"]);

            // Verify that QuotePrefix is set
            Console.WriteLine("A1 QuotePrefix: " + cells["A1"].GetStyle().QuotePrefix);
            Console.WriteLine("A2 QuotePrefix: " + cells["A2"].GetStyle().QuotePrefix);
            Console.WriteLine("B1 QuotePrefix: " + cells["B1"].GetStyle().QuotePrefix);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("QuotePrefixDemo.xlsx");
        }

        // Helper method that creates a style with QuotePrefix = true
        // and applies it to the specified cell using a StyleFlag.
        static void ApplyQuotePrefix(Cell cell)
        {
            // Create a new style and enable QuotePrefix
            Style style = cell.Worksheet.Workbook.CreateStyle();
            style.QuotePrefix = true;

            // Create a StyleFlag that indicates which style attributes to apply
            StyleFlag flag = new StyleFlag();
            flag.QuotePrefix = true;

            // Apply the style to the cell (SetStyle, not ApplyStyle)
            cell.SetStyle(style, flag);
        }
    }
}