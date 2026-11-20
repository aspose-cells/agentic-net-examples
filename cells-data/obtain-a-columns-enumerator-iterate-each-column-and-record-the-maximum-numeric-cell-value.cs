// Title: How to enumerate columns in an Excel worksheet and get the maximum numeric value per column using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# iterate columns and find column max numeric value | C# read Excel file and compute maximum number in each column with Aspose.Cells | Get per‑column maximum numeric cell value from a workbook using Aspose.Cells .NET | How to loop through columns in Aspose.Cells and record highest numeric value | Extract column-wise numeric maxima from an Excel sheet in C#
// Tags: enumerate columns Aspose.Cells | column-wise max numeric value C# | iterate rows and columns Aspose.Cells | extract numeric maxima from Excel worksheet | Aspose.Cells per column max calculation

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet, and enumerates each column up to the last data column. For every column it scans all rows, checks for numeric cells, tracks the highest numeric value, stores the result in a dictionary, and finally prints each column index alongside its maximum numeric value.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found - {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Dictionary to hold maximum numeric value per column (key = column index)
        Dictionary<int, double> columnMaxValues = new Dictionary<int, double>();

        // Iterate through all columns that have data
        int maxColumn = sheet.Cells.MaxDataColumn; // last column with data
        for (int colIndex = 0; colIndex <= maxColumn; colIndex++)
        {
            double maxInColumn = double.MinValue;
            bool hasNumeric = false;

            // Iterate through all rows that have data in the worksheet
            int maxRow = sheet.Cells.MaxDataRow; // last row with data
            for (int row = 0; row <= maxRow; row++)
            {
                // Get the cell at the current row and column
                Cell cell = sheet.Cells[row, colIndex];

                // Check if the cell contains a numeric value
                if (cell.Type == CellValueType.IsNumeric)
                {
                    double val = cell.DoubleValue;
                    if (!hasNumeric || val > maxInColumn)
                    {
                        maxInColumn = val;
                        hasNumeric = true;
                    }
                }
            }

            // Record the maximum if at least one numeric cell was found
            if (hasNumeric)
            {
                columnMaxValues[colIndex] = maxInColumn;
            }
        }

        // Output the results
        foreach (var kvp in columnMaxValues)
        {
            Console.WriteLine($"Column {kvp.Key}: Max Numeric Value = {kvp.Value}");
        }
    }
}
