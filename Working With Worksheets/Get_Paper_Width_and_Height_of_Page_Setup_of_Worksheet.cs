using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class GetPaperSizeDemo
    {
        public static void Run()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the PageSetup object for the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Get the paper width and height in inches (read‑only properties)
            double paperWidthInches = pageSetup.PaperWidth;
            double paperHeightInches = pageSetup.PaperHeight;

            // Display the values
            Console.WriteLine($"Paper Width: {paperWidthInches} inches");
            Console.WriteLine($"Paper Height: {paperHeightInches} inches");

            // Optionally, change the paper size to see different dimensions
            pageSetup.PaperSize = PaperSizeType.PaperLetter; // 8.5 x 11 inches
            Console.WriteLine($"After setting PaperSize to Letter:");
            Console.WriteLine($"Paper Width: {pageSetup.PaperWidth} inches");
            Console.WriteLine($"Paper Height: {pageSetup.PaperHeight} inches");

            // Save the workbook if needed
            workbook.Save("GetPaperSizeDemo.xlsx");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            GetPaperSizeDemo.Run();
        }
    }
}