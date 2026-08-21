// Title: C# – Find and Highlight the Maximum Numeric Cell Using Cells Enumerator in Aspose.Cells
// Description: Demonstrates how to create a workbook, populate data, iterate over all instantiated cells with Cells.GetEnumerator(), detect numeric values via IsNumericValue, track the highest double, apply a yellow background style to that cell, and save the result as an Excel file.
// Keywords: Aspose.Cells C# | Cells.GetEnumerator | maximum numeric cell | highlight cell Aspose | enumerate Excel cells | find max value Aspose.Cells | style cell programmatically | save workbook .NET | Excel automation C#
// Common Searches: Aspose.Cells find max numeric value in worksheet | C# enumerate cells to get highest number | highlight cell with largest value using Aspose.Cells | how to use Cells.GetEnumerator in Aspose.Cells | C# Aspose.Cells example for max value detection
// Developer Intent: Locate the numeric cell with the greatest value, apply visual highlighting, and persist the workbook.
// Use Cases: Identify the top price in a product list and mark it for quick review. | Extract the latest date from a schedule column, color‑code it, and export the sheet. | Determine the peak sensor reading in a data set, style the cell, and generate a report.
// AI Prompts: Generate C# code that scans a worksheet with Aspose.Cells Cells enumerator and returns the address of the cell containing the maximum numeric value. | Show how to highlight the cell with the highest numeric value in yellow and save the workbook to a given path using Aspose.Cells. | Explain how to modify the enumeration loop to ignore DateTime cells when searching for the maximum numeric value.

using System;
using System.Collections;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, populate data, iterate over all instantiated cells with Cells.GetEnumerator(), detect numeric values via IsNumericValue, track the highest double, apply a yellow background style to that cell, and save the result as an Excel file.
    public class FindMaxNumericCell
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample data – you can replace this with loading an existing file
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(12.5);
                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(7);
                cells["A4"].PutValue("Cherry");
                cells["B4"].PutValue(20.3);
                cells["A5"].PutValue("Date");
                cells["B5"].PutValue(15);

                // Variables to keep track of the maximum numeric value and its cell
                double maxValue = double.MinValue;
                Cell maxCell = null;

                // Iterate through all instantiated cells
                IEnumerator enumerator = cells.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    Cell cell = (Cell)enumerator.Current;

                    // Consider only numeric cells (int, double, DateTime)
                    if (cell.IsNumericValue)
                    {
                        double currentValue = cell.DoubleValue; // For DateTime cells this returns OADate, which is also numeric

                        if (currentValue > maxValue)
                        {
                            maxValue = currentValue;
                            maxCell = cell;
                        }
                    }
                }

                // Output the result
                if (maxCell != null)
                {
                    Console.WriteLine($"Maximum numeric value: {maxValue}");
                    Console.WriteLine($"Located at cell: {maxCell.Name}");

                    // Highlight the cell
                    Style style = workbook.CreateStyle();
                    style.ForegroundColor = Color.Yellow;
                    style.Pattern = BackgroundType.Solid;
                    maxCell.SetStyle(style);
                }
                else
                {
                    Console.WriteLine("No numeric cells found in the worksheet.");
                }

                // Save the workbook
                string outputPath = "FindMaxNumericCell.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            FindMaxNumericCell.Run();
        }
    }
}
