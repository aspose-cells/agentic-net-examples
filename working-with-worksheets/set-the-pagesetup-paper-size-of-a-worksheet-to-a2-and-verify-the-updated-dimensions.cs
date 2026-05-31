using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set the paper size of the worksheet to A2
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA2;

            // Verify that the PaperSize property is set to A2
            Console.WriteLine("PaperSize enum value: " + sheet.PageSetup.PaperSize);

            // Retrieve and display the physical dimensions (in inches) of the selected paper size
            double widthInches = sheet.PageSetup.PaperWidth;
            double heightInches = sheet.PageSetup.PaperHeight;
            Console.WriteLine($"Paper Width (inches): {widthInches}");
            Console.WriteLine($"Paper Height (inches): {heightInches}");

            // Save the workbook (lifecycle save)
            workbook.Save("Worksheet_A2_PaperSize.xlsx");
        }
    }
}