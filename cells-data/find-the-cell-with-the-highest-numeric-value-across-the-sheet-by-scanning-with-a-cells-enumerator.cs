// Title: Find and highlight the cell with the highest numeric value using Aspose.Cells C# Cells enumerator
// Description: Creates a workbook, fills it with mixed data, iterates all cells with a Cells enumerator, identifies the numeric cell with the maximum double value, writes the value and address to the console, applies a yellow background to that cell, and saves the file as MaxNumericCell.xlsx.
// Keywords: Aspose.Cells find max numeric cell C# | enumerate worksheet cells Aspose | highlight highest value cell Aspose.Cells | C# maximum numeric value Excel | cell enumerator Aspose.Cells example
// Common Searches: Aspose.Cells C# find maximum numeric cell | how to highlight the largest number in an Excel sheet using Aspose | enumerate all cells to get max value Aspose.Cells | C# code to locate and style the highest numeric entry in a workbook
// Developer Intent: Locate the numeric cell with the greatest value in a worksheet and visually emphasize it.
// Use Cases: Automatically flag the product with the highest sales figure in a generated report. | Identify peak sensor readings in a data log and highlight them before distribution. | Perform a quick data‑quality check by marking the largest numeric entry in a dataset.
// AI Prompts: Generate C# code that uses Aspose.Cells to enumerate every cell, find the maximum numeric value, and return its address. | Show how to apply a yellow background style to the cell containing the highest number in an Aspose.Cells workbook. | Write a script that scans an Excel worksheet with Aspose.Cells, prints the max value and cell reference, and saves the highlighted workbook.

using System;
using System.Collections;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills it with mixed data, iterates all cells with a Cells enumerator, identifies the numeric cell with the maximum double value, writes the value and address to the console, applies a yellow background to that cell, and saves the file as MaxNumericCell.xlsx.
    public class FindMaxNumericCell
    {
        // Entry point required by the project
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data (replace with your own loading logic if needed)
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(2.5);
            cells["A3"].PutValue("Orange");
            cells["B3"].PutValue(1.8);
            cells["A4"].PutValue("Banana");
            cells["B4"].PutValue(3.2);
            cells["C5"].PutValue(10); // another numeric cell

            // Enumerate all cells to find the highest numeric value
            IEnumerator enumerator = cells.GetEnumerator();
            Cell maxCell = null;
            double maxValue = double.MinValue;

            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;

                // Consider only numeric cells (int, double, datetime)
                if (cell.IsNumericValue && cell.Value != null)
                {
                    double current = cell.DoubleValue;
                    if (current > maxValue)
                    {
                        maxValue = current;
                        maxCell = cell;
                    }
                }
            }

            // Output result and optionally highlight the cell
            if (maxCell != null)
            {
                Console.WriteLine($"Maximum numeric value: {maxValue} found at cell {maxCell.Name}");

                // Highlight the cell with a yellow background
                Style highlight = workbook.CreateStyle();
                highlight.ForegroundColor = Color.Yellow;
                highlight.Pattern = BackgroundType.Solid;
                maxCell.SetStyle(highlight);
            }
            else
            {
                Console.WriteLine("No numeric cells were found in the worksheet.");
            }

            // Save the workbook (ensure the directory is writable)
            string outputPath = "MaxNumericCell.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}
