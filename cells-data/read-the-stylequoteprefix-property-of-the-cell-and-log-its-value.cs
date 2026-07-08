using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access a cell and put a value that starts with a single quote
            Cell cell = worksheet.Cells["B10"];
            cell.PutValue("'12345"); // The leading single quote is a quote prefix

            // Retrieve the cell's style and enable the QuotePrefix flag
            Style style = cell.GetStyle();
            style.QuotePrefix = true;
            cell.SetStyle(style);

            // Read the QuotePrefix property and log its value
            bool isQuotePrefixSet = cell.GetStyle().QuotePrefix;
            Console.WriteLine("QuotePrefix is set: " + isQuotePrefixSet);

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("QuotePrefixDemo.xlsx");
        }
    }
}