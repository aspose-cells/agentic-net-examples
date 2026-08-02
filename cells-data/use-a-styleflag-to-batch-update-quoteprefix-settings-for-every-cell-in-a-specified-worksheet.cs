using System;
using Aspose.Cells;

class QuotePrefixBatchUpdate
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // default workbook with one worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data: cells with leading single quote and a normal cell
        worksheet.Cells["A1"].PutValue("'SampleText");
        worksheet.Cells["B2"].PutValue("'12345");
        worksheet.Cells["C3"].PutValue("RegularValue");

        // Create a style and set QuotePrefix to true
        Style quoteStyle = workbook.CreateStyle();
        quoteStyle.QuotePrefix = true;

        // Create a StyleFlag and enable the QuotePrefix flag
        StyleFlag flag = new StyleFlag();
        flag.QuotePrefix = true;

        // Apply the style to every cell in the worksheet using the flag
        worksheet.Cells.ApplyStyle(quoteStyle, flag);

        // Save the workbook
        workbook.Save("QuotePrefixBatchUpdated.xlsx");
    }
}