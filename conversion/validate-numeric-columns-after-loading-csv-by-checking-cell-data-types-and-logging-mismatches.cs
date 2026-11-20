// Title: Validate numeric columns in a CSV using Aspose.Cells for .NET and log non‑numeric values
// Common Searches: asp.net load csv with Aspose.Cells and validate numeric columns | c# check if cell value is numeric after importing CSV using Aspose.Cells | how to log rows with non‑numeric data in specific CSV columns with Aspose.Cells | validate data types of selected columns in a CSV workbook in .NET
// Tags: numeric column validation Aspose.Cells | CSV import cell type checking C# | log non-numeric cells Aspose.Cells | load CSV with LoadOptions Aspose.Cells | verify cell data type workbook .NET

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads a CSV file into an Aspose.Cells Workbook, specifies which column indexes should contain numeric data, iterates through each data row (skipping the header if present), checks each targeted cell's type, and writes a console message for any cell whose value is not numeric.
class Program
{
    static void Main()
    {
        // Path to the CSV file to be validated
        string csvPath = "input.csv";

        // Load the CSV file into a Workbook
        var loadOptions = new LoadOptions(LoadFormat.Csv);
        Workbook workbook = new Workbook(csvPath, loadOptions);
        Worksheet sheet = workbook.Worksheets[0];

        // Define which columns (0‑based index) are expected to contain numeric values
        // Adjust the list according to your CSV structure
        List<int> numericColumns = new List<int> { 0, 2, 4 };

        // If the first row contains headers, start validation from the second row
        int startRow = 1; // change to 0 if there is no header row
        int maxRow = sheet.Cells.MaxDataRow;

        // Iterate through each row and each numeric column to verify data types
        for (int row = startRow; row <= maxRow; row++)
        {
            foreach (int col in numericColumns)
            {
                Cell cell = sheet.Cells[row, col];

                // Skip empty cells
                if (cell.Value == null)
                    continue;

                // Check if the cell type is numeric; if not, log the mismatch
                if (cell.Type != CellValueType.IsNumeric)
                {
                    Console.WriteLine(
                        $"Data type mismatch at Row {row + 1}, Column {col + 1}: " +
                        $"Value \"{cell.StringValue}\" is not numeric.");
                }
            }
        }
    }
}
