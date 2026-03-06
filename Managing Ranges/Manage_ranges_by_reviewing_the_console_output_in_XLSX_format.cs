using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in the worksheet
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Laptop");
            cells["B2"].PutValue(1200.50);
            cells["A3"].PutValue("Phone");
            cells["B3"].PutValue(899.99);
            cells["A4"].PutValue("Tablet");
            cells["B4"].PutValue(450.75);

            // Create a range that covers the data (optional)
            Aspose.Cells.Range dataRange = cells.CreateRange("A1:B4");

            // Export the range to a DataTable (including column names)
            DataTable dt = cells.ExportDataTable(0, 0, 4, 2, true);

            // Output the DataTable content to the console
            Console.WriteLine("Exported DataTable content:");
            foreach (DataRow row in dt.Rows)
            {
                Console.WriteLine($"{row[0]}: {row[1]}");
            }

            // Save the workbook to an XLSX file
            string outputPath = "RangesDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}