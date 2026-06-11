using System;
using Aspose.Cells;

namespace AsposeCellsPrintTitleRowsDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first (active) worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Optional: add some sample data to visualize the effect
            for (int row = 1; row <= 20; row++)
            {
                worksheet.Cells[$"A{row}"].PutValue($"Row {row}");
                worksheet.Cells[$"B{row}"].PutValue($"Data {row}");
            }

            // Set print title rows to repeat rows 1 through 3 on each printed page
            // The property expects an Excel style range string, e.g., "$1:$3"
            worksheet.PageSetup.PrintTitleRows = "$1:$3";

            // Save the workbook (lifecycle save)
            workbook.Save("PrintTitleRowsRows1to3.xlsx");

            Console.WriteLine("Workbook saved with print title rows set to $1:$3.");
        }
    }
}