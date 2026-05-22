using System;
using Aspose.Cells;

namespace AutoFitColumnExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert multiline text into cell A1 (row 0, column 0)
            // Use "\n" for line breaks
            worksheet.Cells["A1"].PutValue("First line of text\nSecond line of text\nThird line of text");

            // Enable text wrapping for the cell so that line breaks are respected
            Style style = worksheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(style);

            // AutoFit the column width for column A (index 0) covering rows 0 to 0
            // This ensures the column expands enough to display the longest line
            worksheet.AutoFitColumn(0, 0, 0);

            // Save the workbook to a file
            workbook.Save("AutoFitColumnMultiline.xlsx");
        }
    }
}