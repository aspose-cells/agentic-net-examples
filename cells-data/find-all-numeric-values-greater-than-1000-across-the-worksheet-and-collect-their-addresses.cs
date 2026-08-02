using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace FindLargeValues
{
    class Program
    {
        static void Main()
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // List to store addresses of cells with numeric values > 1000
            List<string> largeValueAddresses = new List<string>();

            // Iterate through all cells in the worksheet
            foreach (Cell cell in worksheet.Cells)
            {
                // Skip empty cells
                if (cell.Value == null) continue;

                // Check for numeric types (double, int, decimal)
                if (cell.Value is double d && d > 1000)
                {
                    largeValueAddresses.Add(cell.Name);
                }
                else if (cell.Value is int i && i > 1000)
                {
                    largeValueAddresses.Add(cell.Name);
                }
                else if (cell.Value is decimal dec && dec > 1000)
                {
                    largeValueAddresses.Add(cell.Name);
                }
                // If the cell contains a numeric string, try to parse it
                else if (cell.Value is string s && double.TryParse(s, out double parsed) && parsed > 1000)
                {
                    largeValueAddresses.Add(cell.Name);
                }
            }

            // Create a new worksheet to output the results
            int resultSheetIndex = workbook.Worksheets.Add();
            Worksheet resultSheet = workbook.Worksheets[resultSheetIndex];
            resultSheet.Name = "LargeValues";

            // Write the collected addresses to the result sheet
            for (int i = 0; i < largeValueAddresses.Count; i++)
            {
                resultSheet.Cells[i, 0].PutValue(largeValueAddresses[i]);
            }

            // Save the workbook (save rule)
            workbook.Save(outputPath);
        }
    }
}