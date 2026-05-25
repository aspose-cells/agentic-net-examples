using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class DetectApostropheRichTextCells
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (you can iterate over all worksheets if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the used range to limit the iteration
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell has a leading apostrophe (QuotePrefix) and contains rich text
                    bool hasLeadingApostrophe = cell.GetStyle().QuotePrefix;
                    bool isRichText = cell.IsRichText();

                    if (hasLeadingApostrophe && isRichText)
                    {
                        // Log the cell address (e.g., "A1")
                        Console.WriteLine($"Cell with leading apostrophe and rich text found at: {cell.Name}");
                    }
                }
            }

            // Optionally, save the workbook if any modifications were made
            // workbook.Save("output.xlsx");
        }
    }
}