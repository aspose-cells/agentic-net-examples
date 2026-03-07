using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class FitToPagesDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the sheet with enough data to span multiple printed pages
        for (int row = 0; row < 200; row++)
        {
            for (int col = 0; col < 20; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Define the print area that includes all populated cells
        sheet.PageSetup.PrintArea = "A1:T200";

        // Set the worksheet to fit to a specific number of pages when printed
        // Option 1: Use the SetFitToPages method
        sheet.PageSetup.SetFitToPages(2, 3); // 2 pages wide, 3 pages tall

        // Option 2 (equivalent): set the properties directly
        // sheet.PageSetup.FitToPagesWide = 2;
        // sheet.PageSetup.FitToPagesTall = 3;

        // Output the current FitToPages settings for verification
        Console.WriteLine($"FitToPagesWide: {sheet.PageSetup.FitToPagesWide}");
        Console.WriteLine($"FitToPagesTall: {sheet.PageSetup.FitToPagesTall}");
        Console.WriteLine($"IsPercentScale (should be false): {sheet.PageSetup.IsPercentScale}");

        // Save the workbook to an XLSX file
        workbook.Save("FitToPagesDemo.xlsx");
    }
}