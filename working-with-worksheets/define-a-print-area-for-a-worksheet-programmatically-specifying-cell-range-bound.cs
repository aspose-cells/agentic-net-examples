using System;
using Aspose.Cells;

namespace PrintAreaExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data (optional, just for demonstration)
            worksheet.Cells["A1"].PutValue("Header1");
            worksheet.Cells["B1"].PutValue("Header2");
            worksheet.Cells["A2"].PutValue("Data1");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["A3"].PutValue("Data2");
            worksheet.Cells["B3"].PutValue(200);

            // Define the print area (cell range boundaries)
            worksheet.PageSetup.PrintArea = "A1:B3";

            // Save the workbook (lifecycle: save)
            workbook.Save("PrintAreaDemo.xlsx");

            Console.WriteLine("Workbook saved with print area set to A1:B3.");
        }
    }
}