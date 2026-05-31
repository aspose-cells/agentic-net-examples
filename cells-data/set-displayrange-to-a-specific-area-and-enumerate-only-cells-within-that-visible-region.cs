using System;
using System.Collections;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class DisplayRangeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (10 rows × 5 columns)
            for (int r = 0; r < 10; r++)
            {
                for (int c = 0; c < 5; c++)
                {
                    cells[r, c].PutValue($"R{r}C{c}");
                }
            }

            // Define the display range B2:D4 (zero‑based indices: row 1‑3, column 1‑3)
            CellArea displayArea = CellArea.CreateCellArea(1, 1, 3, 3);
            // Note: Setting the view's display range is optional; if required, ensure the Aspose.Cells version supports Worksheet.View.
            // worksheet.View.DisplayRange = displayArea;

            // Build a Range that exactly matches the display area
            int startRow = displayArea.StartRow;
            int startCol = displayArea.StartColumn;
            int rowCount = displayArea.EndRow - displayArea.StartRow + 1;
            int colCount = displayArea.EndColumn - displayArea.StartColumn + 1;
            AsposeRange visibleRange = cells.CreateRange(startRow, startCol, rowCount, colCount);

            // Enumerate only the cells inside the visible range
            IEnumerator enumerator = visibleRange.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }

            // Save the workbook (optional)
            string outputPath = "DisplayRangeDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}