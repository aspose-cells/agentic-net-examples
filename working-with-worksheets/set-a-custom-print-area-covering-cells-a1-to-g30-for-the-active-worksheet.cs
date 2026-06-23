using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first (active) worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the custom print area to cover cells A1:G30
            // This uses the PageSetup.PrintArea property as defined in the API.
            worksheet.PageSetup.PrintArea = "A1:G30";

            // Save the workbook to a file (lifecycle: save)
            // Adjust the file path as needed.
            workbook.Save("PrintAreaDemo.xlsx");

            Console.WriteLine("Print area set to A1:G30 and workbook saved successfully.");
        }
    }
}