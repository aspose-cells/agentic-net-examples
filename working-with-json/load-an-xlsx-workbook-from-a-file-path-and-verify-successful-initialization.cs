using System;
using Aspose.Cells;

namespace AsposeCellsLoadDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string filePath = "example.xlsx";

            // Load the workbook using the string constructor (loads the file)
            Workbook workbook = new Workbook(filePath);

            // Verify that the workbook was initialized successfully
            if (workbook != null && workbook.Worksheets.Count > 0)
            {
                Console.WriteLine("Workbook loaded successfully.");
                Console.WriteLine("Number of worksheets: " + workbook.Worksheets.Count);
            }
            else
            {
                Console.WriteLine("Failed to load the workbook.");
            }
        }
    }
}