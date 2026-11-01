// Title: Remove duplicate values from a named range in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: c# aspocells clear duplicate entries in a named range | how to deduplicate values inside an Excel named range with Aspose.Cells | example code to keep first occurrence and delete repeats in a named range using Aspose.Cells | aspocells remove repeated cells from a specific named range in C# | c# Aspose.Cells tutorial for eliminating duplicate cells in Excel named range
// Tags: Aspose.Cells named range duplicate removal | C# Excel deduplication using Aspose.Cells | Aspose.Cells clear duplicate cells | Excel named range processing with Aspose.Cells | Aspose.Cells workbook save after deduplication

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The program loads an Excel file, retrieves the named range "MyRange", iterates through each cell, and clears any duplicate values while preserving the first occurrence. After processing, the workbook is saved to a new file.
class ReplaceDuplicatesInNamedRange
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the named range "MyRange"
            AsposeRange namedRange = workbook.Worksheets.GetRangeByName("MyRange");
            if (namedRange == null)
            {
                Console.WriteLine("Named range \"MyRange\" not found.");
                return;
            }

            // Determine the boundaries of the named range
            int startRow = namedRange.FirstRow;
            int startColumn = namedRange.FirstColumn;
            int totalRows = namedRange.RowCount;
            int totalColumns = namedRange.ColumnCount;

            // Keep track of values that have already been encountered
            HashSet<string> seenValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Iterate through each cell in the named range
            for (int i = 0; i < totalRows; i++)
            {
                for (int j = 0; j < totalColumns; j++)
                {
                    Cell currentCell = worksheet.Cells[startRow + i, startColumn + j];
                    string cellValue = currentCell.StringValue?.Trim();

                    // Skip empty cells
                    if (string.IsNullOrEmpty(cellValue))
                        continue;

                    // If the value has been seen before, it is a duplicate – clear it
                    if (seenValues.Contains(cellValue))
                    {
                        currentCell.PutValue(string.Empty);
                        continue;
                    }

                    // First occurrence – add to the set
                    seenValues.Add(cellValue);

                    // Clear any further duplicates of this value within the same named range
                    for (int ii = i; ii < totalRows; ii++)
                    {
                        int startJ = (ii == i) ? j + 1 : 0;
                        for (int jj = startJ; jj < totalColumns; jj++)
                        {
                            Cell dupCell = worksheet.Cells[startRow + ii, startColumn + jj];
                            string dupValue = dupCell.StringValue?.Trim();

                            if (!string.IsNullOrEmpty(dupValue) &&
                                string.Equals(dupValue, cellValue, StringComparison.OrdinalIgnoreCase))
                            {
                                dupCell.PutValue(string.Empty);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Processing complete. Output saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
