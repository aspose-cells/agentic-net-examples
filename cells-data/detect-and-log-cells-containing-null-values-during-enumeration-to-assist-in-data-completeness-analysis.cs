using System;
using Aspose.Cells;

namespace AsposeCellsNullDetectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with some empty (null) cells
            worksheet.Cells["A1"].PutValue("Header1");
            worksheet.Cells["B1"].PutValue("Header2");
            worksheet.Cells["A2"].PutValue(100);
            // B2 left empty (null)
            worksheet.Cells["A3"].PutValue(null); // explicit null
            worksheet.Cells["B3"].PutValue(300);

            // Determine the range that contains data
            int maxRow = worksheet.Cells.MaxRow;       // zero‑based index of last row with data
            int maxColumn = worksheet.Cells.MaxColumn; // zero‑based index of last column with data

            // Enumerate each row and column within the data range
            for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++)
            {
                Row row = worksheet.Cells.Rows[rowIndex];

                for (int colIndex = 0; colIndex <= maxColumn; colIndex++)
                {
                    // GetCellOrNull returns null when the cell object does not exist
                    Cell cell = row.GetCellOrNull(colIndex);

                    if (cell == null)
                    {
                        // Cell object is missing – treat as null for completeness analysis
                        Console.WriteLine($"[Missing] Row {rowIndex + 1}, Column {colIndex + 1}");
                    }
                    else if (cell.Value == null)
                    {
                        // Cell exists but its value is null (empty)
                        Console.WriteLine($"[Null] {cell.Name} (Row {rowIndex + 1}, Column {colIndex + 1})");
                    }
                }
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("NullDetectionResult.xlsx");
        }
    }
}