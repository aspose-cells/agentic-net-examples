using System;
using Aspose.Cells;

class FitToPagesDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // (Optional) Populate some data to demonstrate scaling
        for (int i = 0; i < 50; i++)
        {
            sheet.Cells[i, 0].PutValue($"Row {i + 1}");
            sheet.Cells[i, 1].PutValue($"Value {i + 1}");
        }

        // Configure the print settings to fit the worksheet onto 2 pages wide and 3 pages tall
        sheet.PageSetup.SetFitToPages(2, 3);

        // (Optional) Output the applied settings for verification
        Console.WriteLine("FitToPagesWide: " + sheet.PageSetup.FitToPagesWide);
        Console.WriteLine("FitToPagesTall: " + sheet.PageSetup.FitToPagesTall);

        // Save the workbook
        workbook.Save("FitToPagesOutput.xlsx");
    }
}