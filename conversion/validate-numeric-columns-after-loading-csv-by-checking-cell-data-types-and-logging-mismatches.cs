// Title: Validate numeric columns in a CSV with Aspose.Cells for .NET and log type mismatches
// Common Searches: Aspose.Cells C# validate that specific CSV columns contain only numbers | how to detect non-numeric values in CSV columns using Aspose.Cells | C# code to log rows with invalid numeric data after loading CSV with Aspose.Cells | checking cell data type in a CSV worksheet with Aspose.Cells .NET | iterate through CSV rows and verify numeric columns using Aspose.Cells
// Tags: numeric column validation Aspose.Cells | cell type checking CSV Aspose.Cells | log mismatched cell values C# | load CSV workbook Aspose.Cells | iterate worksheet rows Aspose.Cells | specify numeric columns by index Aspose.Cells

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

// The example loads a CSV file into an Aspose.Cells Workbook, defines the column indexes that must contain numeric data, iterates over each row up to the last populated row, checks each targeted cell's Type against CellValueType.IsNumeric, and writes a console message for any cell whose value is not numeric.
class Program
{
    static void Main()
    {
        // Path to the CSV file
        string csvPath = "data.csv";

        // Load the CSV file into a workbook
        var loadOptions = new LoadOptions(LoadFormat.Csv);
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Get the first worksheet (CSV data is loaded into the first sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Define which columns should contain numeric data (0‑based indices)
        // Example: columns A and C are expected to be numeric
        List<int> numericColumns = new List<int> { 0, 2 };

        // Determine the last row with data
        int maxRow = sheet.Cells.MaxDataRow;

        // Iterate through each row and validate numeric columns
        for (int row = 0; row <= maxRow; row++)
        {
            foreach (int col in numericColumns)
            {
                Cell cell = sheet.Cells[row, col];

                // Check if the cell's value type is numeric
                if (cell.Type != CellValueType.IsNumeric)
                {
                    // Log mismatch: row and column numbers are 1‑based for readability
                    Console.WriteLine(
                        $"Mismatch at Row {row + 1}, Column {col + 1}: " +
                        $"Value '{cell.StringValue}' is not numeric.");
                }
            }
        }
    }
}
