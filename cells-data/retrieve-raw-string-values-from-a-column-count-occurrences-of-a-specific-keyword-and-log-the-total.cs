using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Specify the column to inspect (0 = column A) and the keyword to count
        int columnIndex = 0;
        string keyword = "Aspose";

        int occurrenceCount = 0;

        // Iterate through all rows that contain data in the worksheet
        for (int row = 0; row <= cells.MaxDataRow; row++)
        {
            // Retrieve the cell at the current row and specified column
            Cell cell = cells[row, columnIndex];

            // Get the raw string value of the cell (empty string if the cell is null)
            string cellValue = cell?.StringValue ?? string.Empty;

            // Check if the keyword appears in the cell value (case‑insensitive)
            if (!string.IsNullOrEmpty(cellValue) &&
                cellValue.Contains(keyword, StringComparison.OrdinalIgnoreCase))
            {
                occurrenceCount++;
            }
        }

        // Log the total number of occurrences found
        Console.WriteLine($"Total occurrences of \"{keyword}\" in column {CellsHelper.ColumnIndexToName(columnIndex)}: {occurrenceCount}");

        // Save the workbook (optional – here we simply write it back to a new file)
        workbook.Save("output.xlsx");
    }
}