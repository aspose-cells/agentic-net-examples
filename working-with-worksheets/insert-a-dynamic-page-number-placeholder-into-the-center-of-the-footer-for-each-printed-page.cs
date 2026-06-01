using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the center section of the footer to display the current page number
        // Section index: 0 = left, 1 = center, 2 = right
        // Footer script "&P" inserts the current page number
        worksheet.PageSetup.SetFooter(1, "Page &P");

        // Save the workbook to a file
        workbook.Save("PageNumberFooter.xlsx");
    }
}