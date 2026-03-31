using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckAutomaticPaperSize
    {
        public static void Run()
        {
            // Create a new workbook (default workbook)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Determine whether the paper size is set to automatic
            bool isAutomatic = pageSetup.IsAutomaticPaperSize;

            // Display the result
            Console.WriteLine("Is paper size automatic? " + isAutomatic);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            CheckAutomaticPaperSize.Run();
        }
    }
}