using System;
using Aspose.Cells;

namespace AsposeCellsPrintQualityDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the print quality (dots per inch)
            worksheet.PageSetup.PrintQuality = 144;

            // Display the current print quality setting
            Console.WriteLine("Print Quality: " + worksheet.PageSetup.PrintQuality);

            // Save the workbook to a file
            workbook.Save("PrintQualityDemo.xlsx");
        }
    }
}