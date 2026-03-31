using System;
using Aspose.Cells;

namespace CenterPrintDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header");
            for (int row = 2; row <= 10; row++)
            {
                sheet.Cells[$"A{row}"].PutValue($"Item {row - 1}");
                sheet.Cells[$"B{row}"].PutValue(row * 10);
            }

            // Define the range that should be printed
            // This also selects the area to be centered
            sheet.PageSetup.PrintArea = "A1:B10";

            // Center the printed page horizontally and vertically
            sheet.PageSetup.CenterHorizontally = true;
            sheet.PageSetup.CenterVertically = true;

            // Save the workbook (save rule)
            workbook.Save("CenteredPrintDemo.xlsx");

            Console.WriteLine("Workbook saved with centered print settings.");
        }
    }
}