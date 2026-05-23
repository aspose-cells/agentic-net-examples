using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "ProfitMargins"
            Name profitMarginsName = workbook.Worksheets.Names["ProfitMargins"];
            if (profitMarginsName == null)
            {
                Console.WriteLine("Named range 'ProfitMargins' not found.");
                return;
            }

            // Get the actual range object (use alias to avoid ambiguity with System.Range)
            AsposeRange profitMarginsRange = profitMarginsName.GetRange();

            // Iterate through each cell in the range
            for (int row = 0; row < profitMarginsRange.RowCount; row++)
            {
                for (int col = 0; col < profitMarginsRange.ColumnCount; col++)
                {
                    Cell cell = profitMarginsRange[row, col];

                    // Process only numeric cells
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        double currentValue = cell.DoubleValue;

                        // Replace negative values with zero
                        if (currentValue < 0)
                        {
                            cell.PutValue(0);
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}