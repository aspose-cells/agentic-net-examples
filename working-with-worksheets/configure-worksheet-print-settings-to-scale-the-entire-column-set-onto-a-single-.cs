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

        // (Optional) Add some sample data to demonstrate the effect
        worksheet.Cells["A1"].PutValue("Sample Data 1");
        worksheet.Cells["B1"].PutValue("Sample Data 2");
        worksheet.Cells["C1"].PutValue("Sample Data 3");

        // Configure the page setup:
        // Fit all columns onto a single page width.
        // Setting FitToPagesTall to 0 lets the height adjust automatically.
        worksheet.PageSetup.FitToPagesWide = 1; // one page wide
        worksheet.PageSetup.FitToPagesTall = 0; // auto height

        // Save the workbook with the configured print settings
        workbook.Save("FitAllColumnsOnePage.xlsx");
    }
}