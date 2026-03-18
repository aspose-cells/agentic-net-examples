using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Path to the Excel file
        string path = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(path);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Check if the worksheet contains any query tables
        if (worksheet.QueryTables.Count > 0)
        {
            // Get the first query table
            QueryTable queryTable = worksheet.QueryTables[0];

            // Obtain the result range of the query table
            AsposeRange resultRange = queryTable.ResultRange;

            // Use Console.Out (a TextWriter) to stream TXT output to the console
            TextWriter writer = Console.Out;

            // Iterate through each row in the result range
            for (int row = resultRange.FirstRow; row < resultRange.FirstRow + resultRange.RowCount; row++)
            {
                // Iterate through each column in the result range
                for (int col = resultRange.FirstColumn; col < resultRange.FirstColumn + resultRange.ColumnCount; col++)
                {
                    // Retrieve the cell value as a string
                    string cellValue = worksheet.Cells[row, col].StringValue;

                    // Write the cell value, separating columns with a tab character
                    writer.Write(cellValue);
                    if (col < resultRange.FirstColumn + resultRange.ColumnCount - 1)
                        writer.Write("\t");
                }

                // End of row
                writer.WriteLine();
            }
        }
        else
        {
            // No query tables present in the worksheet
            Console.WriteLine("No query tables found in the worksheet.");
        }
    }
}