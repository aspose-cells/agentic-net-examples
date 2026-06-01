using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(95);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(150);

                // Create a named range called "SalesData" that refers to the populated cells
                sheet.Cells.CreateRange("A1", "B4").Name = "SalesData";

                // Retrieve the Range object for the named range "SalesData"
                // Use fully qualified name to avoid conflict with System.Range
                Aspose.Cells.Range salesRange = workbook.Worksheets.GetRangeByName("SalesData");

                // Verify that the range was found
                if (salesRange != null)
                {
                    // Iterate through each cell in the range
                    foreach (Cell cell in salesRange)
                    {
                        Console.WriteLine($"{cell.Name}: {cell.Value}");
                    }
                }
                else
                {
                    Console.WriteLine("Named range 'SalesData' not found.");
                }

                // Save the workbook (optional, just to demonstrate lifecycle compliance)
                string outputPath = "NamedRangeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}