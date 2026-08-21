// Title: C# – Replace negative numbers with zero in the named range “ProfitMargins” using Aspose.Cells
// Description: Loads a workbook, retrieves the named range ProfitMargins, iterates each cell, changes any numeric value below 0 to 0, and saves the modified file. Includes checks for missing files and missing named ranges.
// Keywords: Aspose.Cells C# replace negative values | named range ProfitMargins | set negative to zero | Excel data cleansing Aspose | range iteration .NET
// Common Searches: Aspose.Cells replace negative values in named range | C# set values less than zero to zero in Excel range | How to clean profit margin data with Aspose.Cells | Iterate over named range cells Aspose .NET | Replace negative numbers in Excel using Aspose.Cells
// Developer Intent: Set all numeric cells in the named range ProfitMargins that are less than zero to zero.
// Use Cases: Ensure financial reports never contain negative profit‑margin values before analysis. | Normalize data for charts where negative margins would distort the visual output. | Automate validation in an ETL pipeline that requires non‑negative values in the ProfitMargins range.
// AI Prompts: Write C# code with Aspose.Cells that replaces every negative numeric value with zero in the named range "ProfitMargins" and saves the workbook. | Show how to use Aspose.Cells Range.Replace (or a similar method) to set all values < 0 to 0 within a named range. | Explain error‑handling strategies for missing files, absent named ranges, or non‑numeric cells when replacing negative numbers.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, retrieves the named range ProfitMargins, iterates each cell, changes any numeric value below 0 to 0, and saves the modified file. Includes checks for missing files and missing named ranges.
    public class ReplaceNegativeValuesInNamedRange
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range "ProfitMargins"
                Name profitMarginsName = workbook.Worksheets.Names["ProfitMargins"];
                if (profitMarginsName == null)
                {
                    Console.WriteLine("Named range 'ProfitMargins' does not exist.");
                    return;
                }

                // Get the actual range referenced by the name
                Aspose.Cells.Range profitMarginsRange = profitMarginsName.GetRange();

                // Iterate through each cell in the range and replace negative numbers with zero
                for (int row = 0; row < profitMarginsRange.RowCount; row++)
                {
                    for (int col = 0; col < profitMarginsRange.ColumnCount; col++)
                    {
                        Cell cell = profitMarginsRange[row, col];
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            double value = cell.DoubleValue;
                            if (value < 0)
                            {
                                cell.PutValue(0);
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
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ReplaceNegativeValuesInNamedRange.Run();
        }
    }
}
