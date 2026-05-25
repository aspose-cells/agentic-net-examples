using Aspose.Cells;
using System;

class ApplyQuotePrefixStyleFlag
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access a cell and put a value that should be treated as text
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("12345");

        // Create a style and enable the QuotePrefix property
        Style style = workbook.CreateStyle();
        style.QuotePrefix = true;

        // Create a StyleFlag with only the QuotePrefix flag set to true
        StyleFlag flag = new StyleFlag();
        flag.QuotePrefix = true; // other flags remain false

        // Apply the style to the cell using the flag – only QuotePrefix will be updated
        cell.SetStyle(style, flag);

        // Save the workbook
        workbook.Save("QuotePrefixOnly.xlsx");
    }
}