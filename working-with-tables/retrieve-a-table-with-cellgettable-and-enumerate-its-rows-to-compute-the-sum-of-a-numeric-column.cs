// Title: Compute the sum of a numeric column in an Excel table using Cell.GetTable with Aspose.Cells for .NET
// Common Searches: how to calculate total of a column in an Excel table with Aspose.Cells C# | using Cell.GetTable to iterate rows and sum values in .NET | retrieve ListObject from a cell and sum numeric column Aspose.Cells | Aspose.Cells C# sum values of the "Amount" column in a ListObject | enumerate DataRange rows to compute column total with Aspose.Cells
// Tags: Cell.GetTable ListObject extraction | calculate column total Aspose.Cells | enumerate table DataRange rows | find column index by header Aspose.Cells | save workbook after aggregation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, uses Cell.GetTable to obtain the ListObject that contains cell A1, locates the column named "Amount" (or defaults to the second column), iterates through the table's DataRange rows, parses each cell as a double, accumulates the sum, prints the total, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Anchor cell for the table (adjust if needed)
            Cell anchorCell = worksheet.Cells["A1"];

            // Retrieve the table that contains the anchor cell
            ListObject table = anchorCell.GetTable();

            if (table == null)
            {
                Console.WriteLine("No table found at the specified cell.");
                return;
            }

            // Determine the column index for "Amount" (case‑insensitive)
            int sumColumnIndex = -1;
            for (int i = 0; i < table.ListColumns.Count; i++)
            {
                if (string.Equals(table.ListColumns[i].Name, "Amount", StringComparison.OrdinalIgnoreCase))
                {
                    sumColumnIndex = i;
                    break;
                }
            }

            // Default to the second column if "Amount" not found
            if (sumColumnIndex == -1)
                sumColumnIndex = 1; // zero‑based index within the table

            // Use the table's data range to iterate rows
            AsposeRange dataRange = table.DataRange;
            int firstRow = dataRange.FirstRow;
            int firstColumn = dataRange.FirstColumn;
            int rowCount = dataRange.RowCount;

            double total = 0.0;

            for (int i = 0; i < rowCount; i++)
            {
                // Calculate the cell position for the target column
                int rowIdx = firstRow + i;
                int colIdx = firstColumn + sumColumnIndex;

                object cellValue = worksheet.Cells[rowIdx, colIdx].Value;
                if (cellValue != null && double.TryParse(cellValue.ToString(), out double numeric))
                {
                    total += numeric;
                }
            }

            Console.WriteLine($"Sum of the numeric column: {total}");

            // Save the workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
