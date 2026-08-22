// Title: Identify and count duplicate text values in a specific Excel column using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates a given column, builds a case‑insensitive dictionary of string values, and outputs each duplicate with its occurrence count. | Modify the duplicate‑detection example to apply a yellow fill style to every cell that contains a repeated value. | Create a method that returns a list of row numbers for each duplicate string found in a column using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to list duplicate entries in column A of an Excel file | C# count repeated text values in an Excel worksheet column using Aspose.Cells | detect case‑insensitive duplicate strings in Excel with Aspose.Cells .NET API
// Tags: duplicate detection in Excel column Aspose.Cells | enumerate column values C# Aspose.Cells | case‑insensitive string count Aspose.Cells | highlight duplicate cells Aspose.Cells .NET | retrieve row indices of duplicate entries Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsDuplicateDetection
{
    // The example creates a workbook, fills column A with sample strings, iterates the column to build a case‑insensitive dictionary, records values that appear more than once, prints each duplicate with its occurrence count, and saves the file as DuplicateDetection.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in column A (index 0) with some duplicates
            string[] sampleData = { "Apple", "Banana", "Apple", "Orange", "Banana", "Grape" };
            for (int i = 0; i < sampleData.Length; i++)
            {
                cells[i, 0].PutValue(sampleData[i]);
            }

            // Dictionary to count occurrences of each text value
            Dictionary<string, int> occurrenceMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            // List to keep track of values that are duplicates (appear more than once)
            List<string> duplicateValues = new List<string>();

            // Determine the last row that contains data in the column
            int lastRow = cells.MaxDataRow;

            // Enumerate the column and track occurrences
            for (int row = 0; row <= lastRow; row++)
            {
                string cellValue = cells[row, 0].StringValue;

                // Skip empty cells
                if (string.IsNullOrEmpty(cellValue))
                    continue;

                if (occurrenceMap.ContainsKey(cellValue))
                {
                    occurrenceMap[cellValue]++;

                    // Add to duplicate list only once when the second occurrence is found
                    if (!duplicateValues.Contains(cellValue))
                        duplicateValues.Add(cellValue);
                }
                else
                {
                    occurrenceMap[cellValue] = 1;
                }
            }

            // Output the detected duplicate entries and their counts
            Console.WriteLine("Duplicate entries in column A:");
            foreach (string dup in duplicateValues)
            {
                Console.WriteLine($"{dup} occurs {occurrenceMap[dup]} times");
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("DuplicateDetection.xlsx");
        }
    }
}
