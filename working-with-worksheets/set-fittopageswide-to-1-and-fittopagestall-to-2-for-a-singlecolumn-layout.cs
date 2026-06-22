using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the page setup to fit 1 page wide and 2 pages tall
        worksheet.PageSetup.SetFitToPages(1, 2);
        // Alternatively:
        // worksheet.PageSetup.FitToPagesWide = 1;
        // worksheet.PageSetup.FitToPagesTall = 2;

        // Save the workbook
        workbook.Save("FitToPagesDemo.xlsx");
    }
}