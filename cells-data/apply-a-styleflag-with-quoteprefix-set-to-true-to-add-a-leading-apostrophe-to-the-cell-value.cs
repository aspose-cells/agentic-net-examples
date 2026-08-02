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
            Worksheet worksheet = workbook.Worksheets[0];

            // Access cell A1 and put a value that should be treated as text
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("123456");

            // Create a style and enable the QuotePrefix property
            Style style = workbook.CreateStyle();
            style.QuotePrefix = true;

            // Create a StyleFlag and enable the QuotePrefix flag
            StyleFlag flag = new StyleFlag();
            flag.QuotePrefix = true;

            // Apply the style to the cell using the flag
            cell.SetStyle(style, flag);

            // Save the workbook to verify the effect
            workbook.Save("QuotePrefixDemo.xlsx");
        }
    }
}