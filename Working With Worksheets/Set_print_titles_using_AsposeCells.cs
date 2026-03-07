using System;
using Aspose.Cells;

namespace AsposeCellsPrintTitlesDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            Demo.Run();
        }
    }

    public class Demo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data to illustrate the effect of print titles
            worksheet.Cells["A1"].PutValue("Header Row");
            worksheet.Cells["B1"].PutValue("Header Column");
            for (int i = 2; i <= 30; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Row {i - 1}");
                worksheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Set the rows that will repeat at the top of each printed page (first row)
            worksheet.PageSetup.PrintTitleRows = "$1:$1";

            // Set the columns that will repeat on the left side of each printed page (first column)
            worksheet.PageSetup.PrintTitleColumns = "$A:$A";

            // Save the workbook
            workbook.Save("PrintTitlesDemo.xlsx");
        }
    }
}