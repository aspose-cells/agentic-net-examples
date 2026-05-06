using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello Aspose.Cells!");

            // Save the workbook
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook created successfully.");
        }
    }
}