// Title: How to retrieve an Excel ListObject with Cell.GetTable and sum a numeric column in C# using Aspose.Cells
// Common Searches: C# Aspose.Cells get table from a cell and calculate column total | iterate rows of a ListObject returned by Cell.GetTable in Aspose.Cells | sum values in a specific column of an Excel table using Aspose.Cells .NET | find column index by header name and aggregate numbers with Aspose.Cells | read Excel table DataRange and compute sum of the "Amount" column in C#
// Tags: Cell.GetTable ListObject retrieval | Aspose.Cells column sum calculation | enumerate DataRange rows Aspose.Cells | header based column index Aspose.Cells | numeric aggregation in Excel worksheet C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// Loads Input.xlsx, accesses the first worksheet, uses Cell.GetTable to obtain the ListObject containing cell A1, locates the column named "Amount" by header, iterates over the table's DataRange rows, parses each cell value as a double, accumulates the sum, and prints the result to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get a cell that belongs to the target table (e.g., A1)
            Cell anyCellInTable = sheet.Cells["A1"];

            // Retrieve the table (ListObject) that contains the cell
            ListObject table = anyCellInTable.GetTable();

            if (table == null)
            {
                Console.WriteLine("No table found at the specified cell.");
                return;
            }

            // Determine the index of the numeric column to sum (by column name)
            // Replace "Amount" with the actual header text of the column you want to sum
            string targetHeader = "Amount";
            int columnIndex = -1;
            for (int col = 0; col < table.ListColumns.Count; col++)
            {
                if (string.Equals(table.ListColumns[col].Name, targetHeader, StringComparison.OrdinalIgnoreCase))
                {
                    columnIndex = col;
                    break;
                }
            }

            if (columnIndex == -1)
            {
                Console.WriteLine($"Column \"{targetHeader}\" not found in the table.");
                return;
            }

            // Compute the sum by iterating over the data range of the table
            double sum = 0.0;
            try
            {
                // DataRange includes only the data rows (header excluded)
                AsposeRange dataRange = table.DataRange;
                if (dataRange == null)
                {
                    Console.WriteLine("The table does not contain any data rows.");
                    return;
                }

                int startRow = dataRange.FirstRow;
                int startCol = dataRange.FirstColumn;

                for (int i = 0; i < dataRange.RowCount; i++)
                {
                    Cell cell = sheet.Cells[startRow + i, startCol + columnIndex];
                    object? cellValue = cell?.Value;

                    if (cellValue != null && double.TryParse(cellValue.ToString(), out double numericValue))
                    {
                        sum += numericValue;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while calculating sum: {ex.Message}");
                return;
            }

            // Output the result
            Console.WriteLine($"Sum of column \"{targetHeader}\": {sum}");

            // (Optional) Save the workbook if any modifications were made
            // workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
