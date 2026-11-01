// Title: How to set negative numbers to zero in the "ProfitMargins" named range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# replace negative values with zero in a specific named range | set cells less than 0 to 0 in Excel named range using Aspose.Cells .NET | iterate over ProfitMargins named range and modify numeric cells with Aspose.Cells | C# code to sanitize negative profit margins in an Excel workbook via Aspose.Cells | Aspose.Cells workbook save after correcting negative numbers in a named range
// Tags: zero out negative numbers in named range Aspose.Cells | cell value correction below zero C# Aspose.Cells | iterate over named range cells Aspose.Cells .NET | numeric cell processing in Excel workbook Aspose.Cells | ProfitMargins range adjustment Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, accesses the "ProfitMargins" named range, iterates through its cells, replaces any numeric value less than zero with zero, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "ProfitMargins"
            Aspose.Cells.Range profitMarginsRange = workbook.Worksheets.GetRangeByName("ProfitMargins");
            if (profitMarginsRange == null)
            {
                Console.WriteLine("Named range 'ProfitMargins' not found.");
                return;
            }

            // Get the worksheet that contains the range
            Worksheet ws = profitMarginsRange.Worksheet;

            // Iterate through each cell in the range
            for (int r = 0; r < profitMarginsRange.RowCount; r++)
            {
                for (int c = 0; c < profitMarginsRange.ColumnCount; c++)
                {
                    // Calculate absolute row/column indices
                    int rowIndex = profitMarginsRange.FirstRow + r;
                    int colIndex = profitMarginsRange.FirstColumn + c;

                    Cell cell = ws.Cells[rowIndex, colIndex];

                    // Process only numeric cells
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        double value = cell.DoubleValue;

                        // Replace negative values with zero
                        if (value < 0)
                        {
                            cell.PutValue(0);
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
