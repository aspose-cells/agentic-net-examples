using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsDuplicateDetection
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data for column A (including duplicates)
            string[] sampleData = { "Apple", "Orange", "Apple", "Banana", "Orange", "Grape" };
            for (int i = 0; i < sampleData.Length; i++)
            {
                cells[i, 0].PutValue(sampleData[i]); // Column index 0 = column A
            }

            // Dictionary to track occurrences of each text value
            Dictionary<string, int> occurrenceMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // Determine the last row that contains data in column A
            int lastRow = cells.MaxDataRow;

            // Enumerate the column and count occurrences
            for (int row = 0; row <= lastRow; row++)
            {
                string cellValue = cells[row, 0].StringValue;
                if (string.IsNullOrEmpty(cellValue))
                    continue; // Skip empty cells

                if (occurrenceMap.ContainsKey(cellValue))
                    occurrenceMap[cellValue]++;
                else
                    occurrenceMap[cellValue] = 1;
            }

            // Output duplicate entries
            Console.WriteLine("Duplicate entries in column A:");
            foreach (var kvp in occurrenceMap)
            {
                if (kvp.Value > 1)
                {
                    Console.WriteLine($"{kvp.Key} appears {kvp.Value} times");
                }
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("DuplicateDetection.xlsx");
        }
    }
}