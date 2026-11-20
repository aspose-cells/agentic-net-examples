// Title: Convert string‑based numeric values to true numbers in every cell of an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# convert numeric strings to numbers in entire workbook | How to replace string formatted numbers with actual numeric values using Aspose.Cells .NET | C# example for parsing Excel cells that contain numbers as text with Aspose.Cells | Bulk convert string cells to double across all worksheets in an XLSX file using Aspose.Cells | Change cell type from string to numeric in Excel using Aspose.Cells API C#
// Tags: string-to-number conversion Aspose.Cells | bulk numeric parsing Excel .NET | cell type change PutValue Aspose | process all worksheets numeric conversion | Excel workbook numeric string replacement

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The C# program loads an XLSX workbook with Aspose.Cells, iterates through every worksheet and each cell, detects cells whose type is string but contain a parsable number, converts the string to a double using PutValue, and saves the updated workbook, effectively turning all numeric text into true numeric values.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through each cell in the worksheet
                    foreach (Cell cell in cells)
                    {
                        // If the cell contains a string that can be parsed as a number, convert it to numeric
                        if (cell.Type == CellValueType.IsString)
                        {
                            string str = cell.StringValue;
                            if (double.TryParse(str, out double numericValue))
                            {
                                cell.PutValue(numericValue);
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
