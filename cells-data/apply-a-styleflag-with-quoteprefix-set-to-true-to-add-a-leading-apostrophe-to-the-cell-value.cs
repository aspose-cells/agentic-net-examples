using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the target cell and put a value that should be treated as text
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("123456"); // Without leading apostrophe

            // Create a style and enable the QuotePrefix property
            Style style = workbook.CreateStyle();
            style.QuotePrefix = true; // Indicates the cell value starts with a single quote

            // Create a StyleFlag and enable the QuotePrefix flag
            StyleFlag flag = new StyleFlag();
            flag.QuotePrefix = true; // Apply only the QuotePrefix setting

            // Apply the style to the cell using the flag
            cell.SetStyle(style, flag);

            // Save the workbook to verify the effect
            workbook.Save("QuotePrefixDemo.xlsx");
        }
    }
}