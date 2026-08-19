// Title: Detect duplicate text in an Excel column using Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, fills column A with sample strings, uses a case‑insensitive dictionary to count each value, writes only the duplicated entries with their occurrence count to column B, and saves the file as DuplicateDetectionResult.xlsx.
// Keywords: Aspose.Cells duplicate detection C# | count text occurrences Excel column | find repeated values Aspose.Cells | Excel duplicate summary .NET | dictionary counting Aspose.Cells
// Common Searches: Aspose.Cells how to list duplicate strings in a column | C# count repeated values in Excel with Aspose.Cells | detect duplicate entries in Excel column using .NET | Aspose.Cells sample for duplicate detection
// Developer Intent: Identify text values that appear more than once in a specific worksheet column and report their frequencies.
// Use Cases: Summarize repeated product names in a sales report and display counts next to the original column. | Validate data entry by flagging duplicate IDs and adding a descriptive note beside each duplicate. | Generate an audit sheet that lists inventory items occurring multiple times with occurrence totals.
// AI Prompts: Generate C# code with Aspose.Cells that highlights duplicate cells in column C and adds a comment showing the occurrence count. | Provide an Aspose.Cells example that groups rows by identical text in a column and creates a summary worksheet with duplicate totals. | Show how to modify the duplicate detection sample to write results to a new worksheet instead of column B.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, fills column A with sample strings, uses a case‑insensitive dictionary to count each value, writes only the duplicated entries with their occurrence count to column B, and saves the file as DuplicateDetectionResult.xlsx.
    public class DetectDuplicateTextInColumn
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data in column A (index 0)
                string[] sampleData = { "Apple", "Orange", "Apple", "Banana", "Orange", "Grape", "Apple" };
                for (int i = 0; i < sampleData.Length; i++)
                {
                    cells[i, 0].PutValue(sampleData[i]);
                }

                // Dictionary to track how many times each text appears
                Dictionary<string, int> occurrence = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                // Determine the last row that contains data in the worksheet
                int lastRow = cells.MaxDataRow;

                // Enumerate the cells in column A and count occurrences
                for (int row = 0; row <= lastRow; row++)
                {
                    string value = cells[row, 0].StringValue?.Trim();
                    if (string.IsNullOrEmpty(value))
                        continue; // skip empty cells

                    if (occurrence.ContainsKey(value))
                        occurrence[value]++; // increment existing count
                    else
                        occurrence[value] = 1; // first occurrence
                }

                // Write duplicate detection results to column B
                int outputRow = 0;
                foreach (var kvp in occurrence)
                {
                    if (kvp.Value > 1) // only report duplicates
                    {
                        cells[outputRow, 1].PutValue($"{kvp.Key} appears {kvp.Value} times");
                        outputRow++;
                    }
                }

                // Ensure the output directory exists
                string outputPath = "DuplicateDetectionResult.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (lifecycle rule)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DetectDuplicateTextInColumn.Run();
        }
    }
}
