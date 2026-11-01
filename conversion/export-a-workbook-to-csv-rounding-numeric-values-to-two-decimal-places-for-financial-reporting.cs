// Title: Round numeric cells to two decimal places and export an Excel workbook to CSV using Aspose.Cells for .NET
// Common Searches: C# how to round all numeric cells to two decimals before saving an Excel file as CSV with Aspose.Cells | Aspose.Cells export workbook to CSV with financial rounding of numbers in .NET | Round numbers in each worksheet of an XLSX file and generate CSV using Aspose.Cells C# | Apply Math.Round to numeric cells when converting Excel to CSV with Aspose.Cells | Save first worksheet as CSV after rounding values to two decimal places using Aspose.Cells .NET
// Tags: round numeric cells Aspose.Cells C# | save worksheet as CSV Aspose.Cells | financial rounding Excel to CSV .NET | apply Math.Round to cell values Aspose | process used range before CSV save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an XLSX workbook, iterates through each worksheet’s used range, rounds every numeric cell to two decimal places using Math.Round with MidpointRounding.AwayFromZero, and saves the first worksheet as a CSV file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Get the used range of the worksheet
                AsposeRange usedRange = cells.MaxDisplayRange;
                if (usedRange == null)
                    continue; // Skip empty sheets

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Loop through each cell in the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only numeric cells
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            double originalValue = cell.DoubleValue;
                            double roundedValue = Math.Round(originalValue, 2, MidpointRounding.AwayFromZero);
                            cell.PutValue(roundedValue);
                        }
                    }
                }
            }

            // Save the workbook as CSV (first worksheet will be exported)
            workbook.Save(outputPath, SaveFormat.Csv);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
