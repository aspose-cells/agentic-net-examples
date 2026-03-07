using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class GroupBeforeRangeDemo
    {
        public static void Run()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ==== GROUP MARKER ====
            // Group rows 2 to 4 (zero‑based indices 1‑3) and hide them
            // This operation is performed BEFORE creating the range
            cells.GroupRows(1, 3, true);

            // ==== RANGE BLOCK ====
            // Create a range that starts at cell A1 and spans 5 rows and 5 columns
            // Using the CreateRange rule that takes (firstRow, firstColumn, totalRows, totalColumns)
            AsposeRange dataRange = cells.CreateRange(0, 0, 5, 5);
            dataRange.Name = "MyDataRange";

            // Populate the range with sample values
            for (int i = 0; i < dataRange.RowCount; i++)
            {
                for (int j = 0; j < dataRange.ColumnCount; j++)
                {
                    dataRange[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            GroupBeforeRangeDemo.Run();
        }
    }
}