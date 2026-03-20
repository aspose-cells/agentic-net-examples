using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class PrintTitlesDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Header");
            worksheet.Cells["B1"].PutValue("Value");

            for (int i = 2; i <= 20; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Row {i - 1}");
                worksheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Set the rows to repeat on each printed page (first row)
            worksheet.PageSetup.PrintTitleRows = "$1:$1";

            // Set the columns to repeat on each printed page (first column)
            worksheet.PageSetup.PrintTitleColumns = "$A:$A";

            // Define the print area (optional, includes all data)
            worksheet.PageSetup.PrintArea = "A1:B20";

            // Save the workbook
            workbook.Save("PrintTitlesDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PrintTitlesDemo.Run();
        }
    }
}