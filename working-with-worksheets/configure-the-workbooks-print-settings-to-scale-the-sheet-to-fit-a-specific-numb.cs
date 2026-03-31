using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the PageSetup object of the first worksheet
        PageSetup pageSetup = workbook.Worksheets[0].PageSetup;

        // Scale the sheet to fit 2 pages wide and 3 pages tall when printed
        pageSetup.SetFitToPages(2, 3);

        // Ensure that scaling is based on FitToPages rather than percent scale
        pageSetup.IsPercentScale = false;

        // Save the workbook to a file
        workbook.Save("FitToPagesDemo.xlsx");
    }
}