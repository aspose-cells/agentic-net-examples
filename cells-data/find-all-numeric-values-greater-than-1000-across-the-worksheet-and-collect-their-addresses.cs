using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace FindLargeNumbers
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (or iterate through all worksheets if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // List to hold addresses of cells with numeric values greater than 1000
            List<string> largeValueAddresses = new List<string>();

            // Iterate through all instantiated cells in the worksheet
            foreach (Cell cell in worksheet.Cells)
            {
                // Ensure the cell contains a numeric value
                if (cell.Value is double numericValue)
                {
                    // Check if the value exceeds 1000
                    if (numericValue > 1000)
                    {
                        // Add the cell address (e.g., "B5") to the list
                        largeValueAddresses.Add(cell.Name);
                    }
                }
                else if (cell.Value is int intValue)
                {
                    if (intValue > 1000)
                    {
                        largeValueAddresses.Add(cell.Name);
                    }
                }
                // Add other numeric types if necessary (e.g., decimal, long)
            }

            // Output the collected addresses
            Console.WriteLine("Cells with numeric values > 1000:");
            foreach (string address in largeValueAddresses)
            {
                Console.WriteLine(address);
            }

            // Save the workbook (optional – modify as needed)
            workbook.Save("output.xlsx");
        }
    }
}