using System;
using Aspose.Cells;

namespace AsposeCellsDateConversion
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the last row that contains data
            int lastRow = cells.MaxDataRow;

            // Iterate through each row in column A (index 0)
            for (int row = 0; row <= lastRow; row++)
            {
                Cell sourceCell = cells[row, 0]; // Column A

                // Check if the cell contains a DateTime value
                if (sourceCell.Type == CellValueType.IsDateTime)
                {
                    // Retrieve the DateTime value
                    DateTime dateValue = sourceCell.DateTimeValue;

                    // Convert to ISO 8601 format (e.g., 2023-05-15T13:45:30)
                    string isoString = dateValue.ToString("yyyy-MM-ddTHH:mm:ss");

                    // Store the formatted string in column B (index 1) of the same row
                    cells[row, 1].PutValue(isoString);
                }
                else
                {
                    // If the cell is not a date, you may choose to copy the raw string as is
                    // Uncomment the following line if needed:
                    // cells[row, 1].PutValue(sourceCell.StringValue);
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}