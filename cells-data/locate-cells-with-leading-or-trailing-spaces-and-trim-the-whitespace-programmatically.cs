using System;
using Aspose.Cells;

namespace AsposeCellsTrimWhitespace
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (or iterate through all worksheets if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    // Retrieve the cell
                    Cell cell = cells[row, col];

                    // Process only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;
                        if (!string.IsNullOrEmpty(original))
                        {
                            // Trim leading and trailing whitespace
                            string trimmed = original.Trim();

                            // If trimming changed the value, update the cell
                            if (!original.Equals(trimmed))
                            {
                                cell.PutValue(trimmed);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook (replace with desired output path)
            workbook.Save("output.xlsx");
        }
    }
}