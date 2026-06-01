using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (in-memory)
        Workbook workbook = new Workbook();

        // Access the first worksheet (active worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the PageSetup object for the worksheet
        PageSetup pageSetup = worksheet.PageSetup;

        // Retrieve the paper height (in inches)
        double paperHeight = pageSetup.PaperHeight;

        // Display the paper height in the console
        Console.WriteLine("Paper Height (inches): " + paperHeight);
    }
}