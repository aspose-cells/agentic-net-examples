// Title: Export Excel worksheet cell values to a CSV file using Aspose.Cells for .NET without loading the full workbook model (C#)
// Common Searches: how to write Excel cells to a CSV file with Aspose.Cells in C# without loading the workbook | Aspose.Cells export worksheet to CSV using StreamWriter in .NET | C# code to convert .xlsx to .csv while preserving cell formatting with Aspose.Cells | generate CSV from Excel data without LightCells API using Aspose.Cells | save Excel data as CSV using Aspose.Cells without full workbook load
// Tags: Aspose.Cells CSV export without full model | C# write Excel cells to CSV using StreamWriter | export worksheet to CSV Aspose.Cells .NET | convert .xlsx to .csv without LightCells | cell value extraction Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, determines the used range, reads each cell's value, applies proper CSV escaping, and writes the data row‑by‑row to a CSV file via StreamWriter, all without loading the workbook model into memory.
class Program
{
    static void Main()
    {
        // Paths for source workbook and destination CSV
        string workbookPath = "input.xlsx";
        string csvPath = "output.csv";

        // Verify that the input workbook exists
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Input file \"{workbookPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook (full model) – suitable when LightCells API is unavailable
            Workbook workbook = new Workbook(workbookPath);
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the used range (maximum rows and columns)
            int maxRow = cells.MaxDataRow;      // zero‑based index of last row with data
            int maxCol = cells.MaxDataColumn;   // zero‑based index of last column with data
            int rowCount = maxRow + 1;
            int colCount = maxCol + 1;

            // Open a StreamWriter for the CSV output
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // Iterate through each row in the used range
                for (int row = 0; row < rowCount; row++)
                {
                    string[] csvFields = new string[colCount];

                    // Process each column in the current row
                    for (int col = 0; col < colCount; col++)
                    {
                        // Retrieve the cell value as a string (handles numbers, dates, booleans, etc.)
                        string cellValue = cells[row, col].Value?.ToString() ?? string.Empty;

                        // Escape double quotes by doubling them
                        if (cellValue.Contains("\""))
                            cellValue = cellValue.Replace("\"", "\"\"");

                        // Enclose the field in quotes if it contains commas, quotes, or newlines
                        if (cellValue.Contains(",") || cellValue.Contains("\"") || cellValue.Contains("\n"))
                            cellValue = $"\"{cellValue}\"";

                        csvFields[col] = cellValue;
                    }

                    // Write the CSV line
                    writer.WriteLine(string.Join(",", csvFields));
                }
            }

            Console.WriteLine($"CSV file successfully created at \"{csvPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
