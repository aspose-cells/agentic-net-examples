using System;
using Aspose.Cells;

namespace AsposeCellsRichTextApostropheDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (modify as needed)
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
                    Cell cell = cells[row, col];

                    // Skip empty cells
                    if (cell == null || string.IsNullOrEmpty(cell.StringValue))
                        continue;

                    // Check if the cell starts with a leading apostrophe (QuotePrefix)
                    bool hasLeadingApostrophe = cell.GetStyle().QuotePrefix;

                    // Check if the cell contains rich‑text formatting
                    bool isRichText = cell.IsRichText();

                    // If both conditions are true, log the cell address
                    if (hasLeadingApostrophe && isRichText)
                    {
                        Console.WriteLine($"Cell {cell.Name} (Row {cell.Row + 1}, Column {cell.Column + 1}) contains a leading apostrophe and rich text.");
                    }
                }
            }

            // Optionally save the workbook (if any modifications were made)
            workbook.Save("output.xlsx");
        }
    }
}