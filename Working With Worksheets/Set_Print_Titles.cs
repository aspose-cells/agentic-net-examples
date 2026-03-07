using System;
using Aspose.Cells;

namespace AsposeCellsPrintTitlesDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data to demonstrate the effect of print titles
            worksheet.Cells["A1"].PutValue("Header Row");
            for (int i = 2; i <= 20; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Data Row {i - 1}");
                worksheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Set the rows that will repeat at the top of each printed page (first row)
            worksheet.PageSetup.PrintTitleRows = "$1:$1";

            // Set the columns that will repeat on the left side of each printed page (column A)
            worksheet.PageSetup.PrintTitleColumns = "$A:$A";

            // Save the workbook
            workbook.Save("PrintTitlesDemo.xlsx");
        }
    }
}