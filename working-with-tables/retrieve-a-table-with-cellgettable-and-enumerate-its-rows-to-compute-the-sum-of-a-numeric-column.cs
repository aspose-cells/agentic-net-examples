// Title: How to sum a numeric column in an Excel table using Cell.GetTable with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx workbook, calls Cell.GetTable to obtain the ListObject for a specific cell, locates a column by its header name, and computes the sum of that column's numeric values. | Demonstrate iterating over the DataRange of an Aspose.Cells table to aggregate values from a chosen column while safely handling empty or non‑numeric cells.
// Common Searches: Aspose.Cells C# get table from cell and sum column values | How to calculate total of a column in an Excel ListObject using Aspose.Cells | C# iterate rows of a table retrieved with Cell.GetTable in Aspose.Cells | Sum numeric column named 'Amount' in an Excel table with Aspose.Cells .NET | Retrieve column index by header name in Aspose.Cells ListObject
// Tags: sum column values using Cell.GetTable Aspose.Cells | retrieve ListObject from cell Aspose.Cells C# | iterate table DataRange Aspose.Cells | find column index by header Aspose.Cells ListColumns | calculate numeric column total Excel .xlsx Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// // Loads 'input.xlsx', uses Cell.GetTable to get the ListObject containing cell A1, finds the 'Amount' column index via its header, iterates the table's DataRange rows, parses each cell as a double, accumulates the sum, and prints the result.
class Program
{
    static void Main()
    {
        string filePath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get a cell that belongs to the table (adjust as needed)
            Cell tableCell = worksheet.Cells["A1"];

            // Retrieve the table (ListObject) that contains the cell
            ListObject table = tableCell.GetTable();

            if (table == null)
            {
                Console.WriteLine("No table found at the specified cell.");
                return;
            }

            // Column header to sum
            string numericColumnName = "Amount";

            // Find the column index based on the header name
            int columnIndex = -1;
            for (int i = 0; i < table.ListColumns.Count; i++)
            {
                if (table.ListColumns[i].Name.Equals(numericColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    columnIndex = i;
                    break;
                }
            }

            if (columnIndex == -1)
            {
                Console.WriteLine($"Column \"{numericColumnName}\" not found in the table.");
                return;
            }

            // Get the data range of the table (excluding header/footer)
            AsposeRange dataRange = table.DataRange;
            if (dataRange == null)
            {
                Console.WriteLine("The table does not contain any data rows.");
                return;
            }

            // Compute the sum of the numeric column
            double sum = 0.0;
            int startRow = dataRange.FirstRow;
            int startCol = dataRange.FirstColumn + columnIndex;

            for (int r = 0; r < dataRange.RowCount; r++)
            {
                Cell cell = worksheet.Cells[startRow + r, startCol];
                object cellValue = cell.Value;
                if (cellValue != null && double.TryParse(cellValue.ToString(), out double numericValue))
                {
                    sum += numericValue;
                }
            }

            Console.WriteLine($"Sum of column \"{numericColumnName}\": {sum}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
