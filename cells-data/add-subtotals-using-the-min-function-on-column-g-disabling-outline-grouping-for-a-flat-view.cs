// Title: Insert a flat minimum subtotal row in column G of an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to add a minimum subtotal row to column G with Aspose.Cells in C# without creating an outline | C# Aspose.Cells calculate min value in a column and insert a label row at the end of the sheet | using Aspose.Cells for .NET to write a flat subtotal row after the data range in Excel
// Tags: calculate column minimum Aspose.Cells C# | insert row after data range Aspose.Cells | write label and value to specific cells Aspose.Cells | disable outline grouping flat subtotal Aspose.Cells | Excel worksheet subtotal without grouping .NET

using System;
using System.IO;
using Aspose.Cells;

namespace SubtotalExample
{
    // The example loads an existing workbook, scans column G to find the smallest numeric value, inserts a new row after the last data row, writes the label "Min" in column F and the computed minimum in column G, and saves the file, producing a flat subtotal view without outline grouping.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Determine the data range (assumes the first row is the header)
                int startRow = 0; // Header row
                int endRow = worksheet.Cells.MaxDataRow; // Last row with data

                // Column G corresponds to index 6 (0‑based)
                int columnIndex = 6;

                // Calculate the minimum value in column G
                double minValue = double.MaxValue;
                bool foundNumeric = false;

                for (int row = startRow + 1; row <= endRow; row++)
                {
                    Cell cell = worksheet.Cells[row, columnIndex];
                    if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double val))
                    {
                        if (val < minValue)
                        {
                            minValue = val;
                        }
                        foundNumeric = true;
                    }
                }

                if (foundNumeric)
                {
                    // Insert a new row after the data range to display the subtotal
                    worksheet.Cells.InsertRows(endRow + 1, 1);

                    // Write a label in the preceding column (column F) and the min value in column G
                    worksheet.Cells[endRow + 1, columnIndex - 1].PutValue("Min");
                    worksheet.Cells[endRow + 1, columnIndex].PutValue(minValue);
                }
                else
                {
                    Console.WriteLine("No numeric data found in column G to calculate a minimum.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
