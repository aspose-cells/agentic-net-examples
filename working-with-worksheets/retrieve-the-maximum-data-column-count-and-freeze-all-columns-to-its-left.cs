// Title: How to find the highest used column across all worksheets and freeze columns to its left with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get last column with data across multiple worksheets | freeze panes up to the maximum data column in an Excel workbook using Aspose.Cells | determine maximum used column index in a .xlsx file with Aspose.Cells .NET | apply FreezePanes to all sheets based on the furthest data column in C# | C# script to freeze columns left of the last populated column in every worksheet
// Tags: max used column detection Aspose.Cells | column freeze based on data extent C# | worksheet iteration for column analysis .NET | global FreezePanes application Aspose.Cells | Excel workbook column range calculation

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, scans every worksheet to identify the furthest column that contains data, then freezes all columns to the left of that column on each sheet using Aspose.Cells for .NET, and finally saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Find the maximum data column index across all worksheets
                int maxDataColumnIndex = -1;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int lastCol = sheet.Cells.MaxDataColumn; // zero‑based index of last column with data
                    if (lastCol > maxDataColumnIndex)
                        maxDataColumnIndex = lastCol;
                }

                // Freeze columns up to the maximum data column (if any)
                if (maxDataColumnIndex >= 0)
                {
                    int columnsToFreeze = maxDataColumnIndex + 1; // convert index to count
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Freeze 0 rows and the calculated number of columns
                        sheet.FreezePanes(0, columnsToFreeze, 0, 0);
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
