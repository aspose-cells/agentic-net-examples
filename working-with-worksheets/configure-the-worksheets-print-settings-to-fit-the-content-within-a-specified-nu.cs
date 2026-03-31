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

        // Add sample data to demonstrate multiple pages
        for (int row = 0; row < 50; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Configure the print settings to fit the content to a specific number of pages
        // Here we fit the sheet to 2 pages wide and 3 pages tall
        sheet.PageSetup.SetFitToPages(2, 3);
        // Equivalent alternative:
        // sheet.PageSetup.FitToPagesWide = 2;
        // sheet.PageSetup.FitToPagesTall = 3;

        // Save the workbook with the configured page setup
        workbook.Save("FitToPagesOutput.xlsx");
    }
}