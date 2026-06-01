using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDuplicateDetection
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate column A with sample data containing duplicates
                string[] sampleData = { "Apple", "Orange", "Apple", "Banana", "Orange", "Grape", "Apple" };
                for (int i = 0; i < sampleData.Length; i++)
                {
                    cells[i, 0].PutValue(sampleData[i]); // Column index 0 = "A"
                }

                // ---------- Detect duplicate text entries in column A ----------
                // Dictionary to track occurrence count of each text value
                Dictionary<string, int> occurrenceMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                // Determine the last row that contains data in column A
                // MaxDataRow returns the index of the last row that has any data in the worksheet
                int lastRow = cells.MaxDataRow;

                // Enumerate each cell in the column
                for (int row = 0; row <= lastRow; row++)
                {
                    // Retrieve the cell value as string (trimmed)
                    string cellValue = cells[row, 0].StringValue?.Trim();

                    // Skip empty cells
                    if (string.IsNullOrEmpty(cellValue))
                        continue;

                    // Update occurrence count
                    if (occurrenceMap.ContainsKey(cellValue))
                        occurrenceMap[cellValue]++;
                    else
                        occurrenceMap[cellValue] = 1;
                }

                // Output duplicate entries and their counts
                Console.WriteLine("Duplicate entries in column A:");
                foreach (var kvp in occurrenceMap)
                {
                    if (kvp.Value > 1)
                    {
                        Console.WriteLine($"Value \"{kvp.Key}\" occurs {kvp.Value} times.");
                    }
                }

                // ---------- Save the workbook ----------
                string outputPath = "DuplicateDetectionResult.xlsx";

                // Ensure the directory exists (optional safety)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}