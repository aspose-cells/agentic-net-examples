// Title: Calculate the median of a cell range with Aspose.Cells for .NET and store the result in another cell (C#)
// AI Prompts: Generate C# code using Aspose.Cells that fills cells A1:A10 with numeric values, sets cell B1 formula to MEDIAN(A1:A10), forces formula calculation, and prints the median to the console. | Show how to programmatically evaluate an Excel MEDIAN function in a worksheet with Aspose.Cells, retrieve the computed value, and save the workbook to a file. | Provide a try‑catch example that demonstrates populating a column from an array, applying the MEDIAN formula, calling Workbook.CalculateFormula, and handling potential exceptions.
// Common Searches: Aspose.Cells C# example for MEDIAN function on a range | How to use Aspose.Cells to calculate median of column A and write result to B1 | Programmatically evaluate Excel formulas with Aspose.Cells .NET | Saving workbook after formula calculation using Aspose.Cells C# | Retrieve calculated value of MEDIAN(A1:A10) with Aspose.Cells API
// Tags: Aspose.Cells MEDIAN formula C# | populate worksheet cells from array Aspose.Cells | calculate range median Aspose.Cells | evaluate Excel formulas programmatically .NET | save workbook after formula calculation Aspose.Cells | exception handling Aspose.Cells workbook operations

using System;
using Aspose.Cells;

namespace AsposeCellsMedianExample
{
    // The example creates a new Workbook, fills cells A1‑A10 with numeric data, assigns the MEDIAN(A1:A10) formula to cell B1, forces calculation with Workbook.CalculateFormula, outputs the median value to the console, and saves the file as MedianExample.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data in A1:A10.
                double[] sampleData = { 5, 3, 8, 12, 7, 9, 2, 4, 6, 10 };
                for (int i = 0; i < sampleData.Length; i++)
                {
                    sheet.Cells[i, 0].PutValue(sampleData[i]); // Column A (index 0)
                }

                // Use the built‑in MEDIAN function in cell B1.
                sheet.Cells["B1"].Formula = "MEDIAN(A1:A10)";

                // Calculate the workbook to evaluate the formula.
                workbook.CalculateFormula();

                // Output the result to the console.
                Console.WriteLine("Median of A1:A10 = " + sheet.Cells["B1"].Value);

                // Save the workbook (optional).
                string outputPath = "MedianExample.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
