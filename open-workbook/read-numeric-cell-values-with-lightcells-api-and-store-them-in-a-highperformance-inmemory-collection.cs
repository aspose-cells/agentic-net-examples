// Title: How to read only numeric cells from an Excel worksheet with Aspose.Cells in C# and store them in a high‑performance in‑memory collection
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, iterates over the used range of the first worksheet, and adds each numeric cell value to a List<double>. | Adapt the sample to use a thread‑safe collection such as ConcurrentBag<double> for storing numeric values when processing the worksheet in parallel. | Create a C# snippet that calculates the total count and sum of the numeric values collected from the worksheet and prints the results.
// Common Searches: aspnet read numeric values from Excel using Aspose.Cells without loading entire workbook | c# extract double values from first worksheet cells with Aspose.Cells | high‑performance in‑memory collection for Excel numeric data .net core | how to count and sum numeric cells in an .xlsx file using Aspose.Cells C# | using Aspose.Cells to filter numeric cells only in a worksheet
// Tags: Aspose.Cells numeric cell extraction C# | read used cells performance .NET | store Excel double values in List<double> | thread‑safe collection for Excel data Aspose | calculate sum of numeric cells Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example opens an .xlsx file with Aspose.Cells, accesses the first worksheet, iterates through all used cells, selects those whose type is numeric, adds their double values to a List<double>, and then outputs the total number of numeric cells and their sum, with file‑existence checking and exception handling.
public class Program
{
    public static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Collection to store numeric values
            List<double> numericValues = new List<double>();

            // Iterate through all used cells in the worksheet
            foreach (Cell cell in worksheet.Cells)
            {
                // Check if the cell contains a numeric value
                if (cell.Type == CellValueType.IsNumeric)
                {
                    numericValues.Add(cell.DoubleValue);
                }
            }

            // Output results
            Console.WriteLine($"Total numeric cells read: {numericValues.Count}");

            double sum = 0;
            foreach (double val in numericValues)
            {
                sum += val;
            }
            Console.WriteLine($"Sum of numeric values: {sum}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
