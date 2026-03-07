using System;
using Aspose.Cells;

class RemoveSpecificPageBreak
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some horizontal page breaks
        worksheet.HorizontalPageBreaks.Add(5);   // Row 5
        worksheet.HorizontalPageBreaks.Add(10);  // Row 10
        worksheet.HorizontalPageBreaks.Add(15);  // Row 15

        // Remove the horizontal page break at index 1 (row 10)
        worksheet.HorizontalPageBreaks.RemoveAt(1);

        // Add some vertical page breaks
        worksheet.VerticalPageBreaks.Add(3);   // Column 3
        worksheet.VerticalPageBreaks.Add(7);   // Column 7

        // Remove the vertical page break at index 0 (column 3)
        worksheet.VerticalPageBreaks.RemoveAt(0);

        // Save the workbook
        workbook.Save("RemovedPageBreaks.xlsx");
    }
}