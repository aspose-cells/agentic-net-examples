// Title: C# Aspose.Cells example to group rows by category column and collapse them for a compact worksheet view
// AI Prompts: Generate C# code using Aspose.Cells that reads an Excel file, groups consecutive rows with identical values in column A, and collapses each group. | Write a C# console program that loads a workbook, detects changes in the first column, applies row grouping with hidden detail rows, and saves the updated file. | Adapt an existing Aspose.Cells script to automatically hide the rows after grouping them based on a category column.
// Common Searches: Aspose.Cells C# how to auto‑group rows based on column value | C# collapse grouped rows in an Excel workbook using Aspose.Cells | Programmatically hide row groups in a .xlsx file with Aspose.Cells for .NET | Create a compact view of Excel data by grouping rows with Aspose.Cells C# example | How to group rows by category and hide details in Excel using Aspose.Cells
// Tags: Aspose.Cells row grouping by first column | collapse row groups in Excel with Aspose.Cells | auto hide grouped rows .NET Aspose.Cells | category column grouping Aspose.Cells | compact worksheet view Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsGroupingExample
{
    // The sample loads an existing Excel workbook, scans column A for category changes, groups consecutive rows that share the same category, collapses each group to produce a compact view, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (can be supplied via command‑line arguments)
            string inputFilePath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputFilePath = args.Length > 1 ? args[1] : "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFilePath))
                {
                    Console.WriteLine($"Input file not found: {inputFilePath}");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputFilePath);
                var worksheet = workbook.Worksheets[0];

                // Assume the first row (index 0) contains headers; data starts at row 1
                int startDataRow = 1;
                int lastDataRow = worksheet.Cells.MaxDataRow;
                int categoryColumnIndex = 0; // Column A

                // No data to process
                if (lastDataRow < startDataRow)
                {
                    Console.WriteLine("No data rows found in the worksheet.");
                    workbook.Save(outputFilePath);
                    return;
                }

                // Initialize grouping variables
                int groupStartRow = startDataRow;
                string previousCategory = worksheet.Cells[groupStartRow, categoryColumnIndex].StringValue;

                // Iterate through data rows to detect category changes
                for (int row = startDataRow + 1; row <= lastDataRow; row++)
                {
                    string currentCategory = worksheet.Cells[row, categoryColumnIndex].StringValue;

                    if (!currentCategory.Equals(previousCategory, StringComparison.Ordinal))
                    {
                        // If the previous group has more than one row, collapse it
                        if (row - 1 > groupStartRow)
                        {
                            worksheet.Cells.GroupRows(groupStartRow, row - 1, false);
                        }

                        // Start a new group
                        groupStartRow = row;
                        previousCategory = currentCategory;
                    }
                }

                // Handle the final group
                if (lastDataRow > groupStartRow)
                {
                    worksheet.Cells.GroupRows(groupStartRow, lastDataRow, false);
                }

                // Save the modified workbook
                workbook.Save(outputFilePath);
                Console.WriteLine($"Workbook saved successfully to: {outputFilePath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
