using System;
using Aspose.Cells;

namespace AsposeCellsSaveExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default format is XLSX)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Modify some cell values
            worksheet.Cells["A1"].PutValue("Hello");
            worksheet.Cells["B1"].PutValue("World");

            // Save the workbook back to XLSX format using the default Save method
            workbook.Save("ModifiedWorkbook.xlsx");

            // Optional: inform that the file has been saved
            Console.WriteLine("Workbook saved as ModifiedWorkbook.xlsx");
        }
    }
}