using System;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook in memory
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a value in cell A1
            sheet.Cells["A1"].PutValue("Sample Value");

            // Output the name of the first worksheet and the value of cell A1
            Console.WriteLine($"First worksheet name: {sheet.Name}");
            Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].StringValue}");
        }
    }
}