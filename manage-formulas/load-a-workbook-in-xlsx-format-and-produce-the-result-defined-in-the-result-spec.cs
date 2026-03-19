using System;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLSX file to be loaded
            string filePath = "input.xlsx";

            // Load the workbook using the constructor that accepts a file path
            Workbook workbook = new Workbook(filePath);

            // Output basic information about the loaded workbook
            Console.WriteLine($"Workbook loaded from: {filePath}");
            Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");

            // Iterate through each worksheet and display its name
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                Console.WriteLine($"Worksheet {i + 1}: {sheet.Name}");

                // Example: display the value of cell A1 if it exists
                Cell cellA1 = sheet.Cells["A1"];
                if (cellA1.Value != null)
                {
                    Console.WriteLine($"  A1 Value: {cellA1.Value}");
                }
            }

            // No saving is required for this task; the workbook is only read.
        }
    }
}