// Title: Import Decimal ArrayList and Round Cells to Two Decimals with Aspose.Cells for .NET (C#)
// Description: Creates a Workbook, imports a vertical ArrayList of decimal values into column A using Cells.ImportArrayList, loops through the imported rows, rounds each numeric cell to two decimal places with Math.Round, writes the rounded value back via Cell.PutValue, and saves the result as an XLSX file.
// Keywords: Aspose.Cells | ImportArrayList | C# decimal import | round Excel cells | two decimal places | vertical list import | Cell.PutValue | Math.Round | Excel .NET | financial data formatting
// Common Searches: Aspose.Cells import ArrayList C# | round numbers after import Aspose.Cells | set two decimal places in Excel using Aspose.Cells | C# import decimal list to Excel | cells.ImportArrayList vertical list
// Developer Intent: Load a collection of decimal values into an Excel worksheet and round each entry to two decimal places using Aspose.Cells in C#.
// Use Cases: Export financial figures stored in an ArrayList to Excel and ensure every amount displays with two‑decimal precision for reporting. | Import sensor or measurement data collected in C# collections, round each reading to two decimals, and generate a clean summary workbook. | Create a template that receives raw decimal inputs, standardizes them by rounding, and saves the sheet for downstream analysis or distribution.
// AI Prompts: Generate C# code that uses Aspose.Cells to import an ArrayList of decimal numbers vertically starting at A1 and rounds each cell to two decimal places before saving. | Show how to iterate over a range imported with Cells.ImportArrayList, apply Math.Round to decimal, double, int, or float values, and update the cells with Cell.PutValue. | Explain how to combine Cells.ImportArrayList with a number format string to display two decimal places automatically after rounding.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a Workbook, imports a vertical ArrayList of decimal values into column A using Cells.ImportArrayList, loops through the imported rows, rounds each numeric cell to two decimal places with Math.Round, writes the rounded value back via Cell.PutValue, and saves the result as an XLSX file.
    public class ImportArrayListAndRoundDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet's cells collection
                Workbook workbook = new Workbook(); // create
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Prepare an ArrayList of decimal numbers
                ArrayList decimalData = new ArrayList
                {
                    123.4567m,
                    89.1234m,
                    45.6789m,
                    0.9999m,
                    100.0m
                };

                // Import the ArrayList vertically starting at cell A1 (row 0, column 0)
                cells.ImportArrayList(decimalData, 0, 0, true);

                // Round each imported cell value to two decimal places
                for (int row = 0; row < decimalData.Count; row++)
                {
                    Cell cell = cells[row, 0];
                    if (cell.Value != null && (cell.Value is double || cell.Value is decimal || cell.Value is int || cell.Value is float))
                    {
                        decimal original = Convert.ToDecimal(cell.Value);
                        decimal rounded = Math.Round(original, 2);
                        cell.PutValue(rounded);
                    }
                }

                // Save the workbook
                string outputPath = "ImportArrayListRounded.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ImportArrayListAndRoundDemo.Run();
        }
    }
}
