using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("Item1");
            worksheet.Cells["B2"].PutValue(2);
            worksheet.Cells["A3"].PutValue("Item2");
            worksheet.Cells["B3"].PutValue(3);

            // Set the print area to the desired range
            worksheet.PageSetup.PrintArea = "A1:B3";

            // Save the workbook (Excel format)
            workbook.Save("PrintAreaDemo.xlsx");

            Console.WriteLine("Workbook saved with print area set.");
        }
    }
}