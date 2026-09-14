// Title: Import a decimal ArrayList into an Excel worksheet and round each cell to two decimal places with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to import a System.Collections.ArrayList of decimal values vertically into column A of a new workbook, round every imported cell to two decimal places, and save the file as an .xlsx document. | Show how to adjust the import routine to place the decimal values horizontally across the first row, apply a numeric format with two decimal places, and save the workbook using Aspose.Cells. | Create C# logic that iterates over cells after importing an ArrayList with Aspose.Cells, skips non‑numeric entries, rounds numeric values to two decimal places, and writes the results back to the worksheet.
// Common Searches: c# aspocells import arraylist of decimals and round to two decimal places | how to round imported numeric cells in Aspose.Cells .NET | vertical import of decimal values into Excel using Aspose.Cells | Aspose.Cells round cells after ImportArrayList method | convert decimal to double and round in Aspose.Cells workbook
// Tags: ImportArrayList decimal values Aspose.Cells | round cell values two decimal places .NET | vertical data import Aspose.Cells C# | numeric formatting after import Aspose.Cells | handle decimal and double types Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, importing a System.Collections.ArrayList of decimal numbers vertically into column A, rounding each imported cell to two decimal places, and saving the workbook as ImportArrayListRounded.xlsx using Aspose.Cells for .NET.
    public class ImportArrayListAndRoundDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet's cells collection
                Workbook workbook = new Workbook();
                Cells cells = workbook.Worksheets[0].Cells;

                // Prepare an ArrayList of decimal numbers
                ArrayList decimalData = new ArrayList
                {
                    12.3456m,
                    78.9012m,
                    3.14159m,
                    0.9999m,
                    123.4567m
                };

                // Import the ArrayList vertically starting at cell A1 (row 0, column 0)
                // Parameters: (ArrayList, firstRow, firstColumn, isVertical)
                cells.ImportArrayList(decimalData, 0, 0, true);

                // Round each imported cell value to two decimal places
                for (int row = 0; row < decimalData.Count; row++)
                {
                    // Retrieve the cell that was just populated
                    Cell cell = cells[row, 0];

                    // The ImportArrayList method stores numbers as double by default,
                    // so we handle both decimal and double types.
                    if (cell.Value is decimal decVal)
                    {
                        cell.PutValue(Math.Round(decVal, 2));
                    }
                    else if (cell.Value is double dblVal)
                    {
                        cell.PutValue(Math.Round(dblVal, 2));
                    }
                    else if (cell.Value != null)
                    {
                        // Fallback: try to convert to double and round
                        if (double.TryParse(cell.Value.ToString(), out double parsed))
                        {
                            cell.PutValue(Math.Round(parsed, 2));
                        }
                    }
                }

                // Save the workbook to a file
                workbook.Save("ImportArrayListRounded.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ImportArrayListAndRoundDemo.Run();
        }
    }
}
