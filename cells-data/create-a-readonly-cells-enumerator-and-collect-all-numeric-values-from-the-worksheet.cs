// Title: Read‑only enumeration of worksheet cells with Aspose.Cells for .NET to collect all numeric values
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates the worksheet's Cells collection in read‑only mode, and adds every numeric cell (double, int, decimal) to a List<double>. | Update the enumeration to skip empty cells, catch exceptions per cell, and log the cell address when a conversion error occurs. | Refactor the example to return the collected numeric values as a double[] array instead of printing them to the console.
// Common Searches: Aspose.Cells C# read‑only cell enumeration without altering worksheet | How to extract all numbers from an Excel sheet using Aspose.Cells for .NET | Collect double and integer values from worksheet cells with Aspose.Cells | Skip null cells and log errors while iterating cells in Aspose.Cells C#
// Tags: cells iteration without modification Aspose.Cells | extract numeric values from Excel worksheet | collect double int decimal cells C# | cell-level exception logging Aspose.Cells | numeric values list<double> Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads an Excel file using Aspose.Cells, iterates the first worksheet's Cells collection in a read‑only manner, gathers every double, int, or decimal value into a List<double>, and logs any cell‑level errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // List to store all numeric values found in the worksheet
        List<double> numericValues = new List<double>();

        // Iterate through all cells that contain data
        foreach (Cell cell in sheet.Cells)
        {
            try
            {
                if (cell.Value == null) continue;

                // Directly handle double, int, and decimal types
                if (cell.Value is double d)
                {
                    numericValues.Add(d);
                }
                else if (cell.Value is int i)
                {
                    numericValues.Add(Convert.ToDouble(i));
                }
                else if (cell.Value is decimal dec)
                {
                    numericValues.Add(Convert.ToDouble(dec));
                }
            }
            catch (Exception ex)
            {
                // Log and continue on cell‑level errors
                Console.WriteLine($"Error processing cell {cell.Name}: {ex.Message}");
            }
        }

        // Output the collected numeric values
        Console.WriteLine($"Numeric values found: {numericValues.Count}");
        foreach (double val in numericValues)
        {
            Console.WriteLine(val);
        }

        // (Optional) Save the workbook if any changes were made elsewhere
        // workbook.Save("output.xlsx");
    }
}
