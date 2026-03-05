using System;
using Aspose.Cells;

namespace AsposeCellsConsoleDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing XLSX file
            string filePath = "input.xlsx";

            // Load the workbook using the constructor that accepts a file path
            Workbook workbook = new Workbook(filePath);

            // Output basic workbook information to the console
            Console.WriteLine($"Workbook loaded from: {filePath}");
            Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");

            // Iterate through each worksheet and display its name and the value of cell A1 (if any)
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                Console.WriteLine($"Worksheet {i + 1}: {sheet.Name}");

                // Retrieve the value of cell A1
                Cell cellA1 = sheet.Cells["A1"];
                if (cellA1.Value != null)
                {
                    Console.WriteLine($"  A1 = {cellA1.Value}");
                }
                else
                {
                    Console.WriteLine("  A1 is empty.");
                }
            }

            // Dispose the workbook when done
            workbook.Dispose();
        }
    }
}