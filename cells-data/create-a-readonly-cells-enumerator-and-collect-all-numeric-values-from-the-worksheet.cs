// Title: Iterate through worksheet cells with a foreach loop and extract all numeric values using Aspose.Cells for .NET
// Common Searches: how to loop through all cells in an Excel worksheet using Aspose.Cells C# | extract numeric data from every cell in a workbook with Aspose.Cells | C# foreach cell enumeration example for Aspose.Cells | collect double, int, and decimal values from Excel using Aspose.Cells API | load an XLSX file and read numeric cells without modifying the workbook in C# | Aspose.Cells read‑only style cell iteration to gather numbers
// Tags: Aspose.Cells cell collection iteration | numeric value extraction from Excel cells C# | load workbook with LoadOptions Xlsx Aspose.Cells | read‑only mode cell enumeration Aspose.Cells | aggregate numeric types from Excel cells Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an XLSX workbook with LoadOptions, accesses the first worksheet, iterates over its Cells collection using a foreach loop, checks each cell's Value for numeric types (double, int, decimal, float), adds the numbers to a List<double>, and prints the total count of numeric entries found.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";

                // Prevent FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (read‑only mode not supported in this version)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Collect all numeric values
                List<double> numericValues = new List<double>();

                // Iterate over all cells using foreach
                foreach (Cell cell in worksheet.Cells)
                {
                    if (cell.Value != null)
                    {
                        switch (cell.Value)
                        {
                            case double d:
                                numericValues.Add(d);
                                break;
                            case int i:
                                numericValues.Add(i);
                                break;
                            case decimal dec:
                                numericValues.Add((double)dec);
                                break;
                            case float f:
                                numericValues.Add(f);
                                break;
                            // Extend with other numeric types if needed
                        }
                    }
                }

                Console.WriteLine($"Total numeric values found: {numericValues.Count}");
                // Example: print values
                // foreach (double val in numericValues) Console.WriteLine(val);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
