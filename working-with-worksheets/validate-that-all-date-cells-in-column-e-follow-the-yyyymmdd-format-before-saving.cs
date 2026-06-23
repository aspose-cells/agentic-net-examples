using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsDateValidation
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the last row that contains data
            int lastRow = worksheet.Cells.MaxDataRow;

            // Column E has index 4 (zero‑based)
            const int dateColumnIndex = 4;

            // Iterate through each used row in column E
            for (int row = 0; row <= lastRow; row++)
            {
                Cell cell = worksheet.Cells[row, dateColumnIndex];

                // Skip empty cells
                if (cell == null || string.IsNullOrWhiteSpace(cell.StringValue))
                    continue;

                string cellText = cell.StringValue.Trim();

                // Validate the date format yyyy‑MM‑dd
                bool isValid = DateTime.TryParseExact(
                    cellText,
                    "yyyy-MM-dd",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out _);

                if (!isValid)
                {
                    // If any cell does not match the required format, abort saving
                    throw new InvalidOperationException(
                        $"Cell {cell.Name} contains an invalid date format: \"{cellText}\". Expected format is yyyy-MM-dd.");
                }
            }

            // All dates are valid; save the workbook
            workbook.Save("output.xlsx");
        }
    }
}