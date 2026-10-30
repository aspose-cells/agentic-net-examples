// Title: How to extract a two‑dimensional range returned by a custom Excel function using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get array result from custom worksheet function after CalculateFormula | read multi‑cell range returned by user‑defined function in Aspose.Cells | retrieve object[,] from evaluated cell B2 using Aspose.Cells .NET | extract dimensions of custom function range result in C# Aspose.Cells | handle custom function returning a range in Aspose.Cells workbook
// Tags: custom function range extraction Aspose.Cells | evaluate formula object array Aspose.Cells | read two dimensional array from cell Aspose.Cells | C# retrieve multi‑cell range Aspose.Cells | Aspose.Cells calculate custom functions

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook with Aspose.Cells, forces calculation of all formulas (including custom functions), accesses cell B2, checks whether the evaluated value is a two‑dimensional object array, prints the range size and each element, and saves the workbook.
    class Program
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook; set a password here if the file is protected
                LoadOptions loadOptions = new LoadOptions();
                // loadOptions.Password = "yourPassword"; // uncomment and set if needed

                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Force calculation of all formulas, including custom functions
                workbook.CalculateFormula();

                // Access the worksheet and the cell that contains the custom function
                Worksheet sheet = workbook.Worksheets[0];
                Cell funcCell = sheet.Cells["B2"]; // adjust the address as needed

                // Retrieve the evaluated result of the custom function
                object result = funcCell.Value;

                // If the custom function returns a range, it is provided as a two‑dimensional object array
                if (result is object[,] range)
                {
                    int rows = range.GetLength(0);
                    int cols = range.GetLength(1);
                    Console.WriteLine($"Custom function returned a range of size {rows}x{cols}:");

                    for (int i = 0; i < rows; i++)
                    {
                        for (int j = 0; j < cols; j++)
                        {
                            Console.Write($"{range[i, j]}\t");
                        }
                        Console.WriteLine();
                    }
                }
                else
                {
                    // The function returned a single value or an unexpected type
                    Console.WriteLine($"Custom function did not return a range. Value: {result}");
                }

                // Save the workbook if any changes were made
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
