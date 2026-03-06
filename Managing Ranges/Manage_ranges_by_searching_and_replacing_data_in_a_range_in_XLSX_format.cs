using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the range where the replace operation should be performed
        // Example: cells B2 to D5
        AsposeRange range = worksheet.Cells.CreateRange("B2", "D5");

        // Define the text to find and its replacement
        string placeholder = "Apple";
        string newValue = "Mango";

        // Iterate through each cell in the defined range
        for (int i = 0; i < range.RowCount; i++)
        {
            for (int j = 0; j < range.ColumnCount; j++)
            {
                // Get the current cell
                var cell = range[i, j];

                // If the cell contains the placeholder string, replace it
                if (cell.Type == CellValueType.IsString && cell.StringValue == placeholder)
                {
                    cell.PutValue(newValue);
                }
            }
        }

        // Save the workbook with the changes
        workbook.Save("output.xlsx");
    }
}