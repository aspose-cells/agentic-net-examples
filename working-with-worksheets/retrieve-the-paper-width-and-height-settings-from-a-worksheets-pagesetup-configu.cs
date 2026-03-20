using System;
using Aspose.Cells;

class RetrievePaperSize
{
    static void Main()
    {
        // Create a new workbook (in-memory)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the PageSetup object
        PageSetup pageSetup = worksheet.PageSetup;

        // Retrieve paper width and height in inches
        double paperWidth = pageSetup.PaperWidth;
        double paperHeight = pageSetup.PaperHeight;

        // Display the values
        Console.WriteLine($"Paper Width (inches): {paperWidth}");
        Console.WriteLine($"Paper Height (inches): {paperHeight}");

        // Save the workbook (optional)
        workbook.Save("PaperSizeDemo.xlsx");
    }
}