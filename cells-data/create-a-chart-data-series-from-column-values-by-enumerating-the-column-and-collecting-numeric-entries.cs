using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesFromColumn
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column A with mixed data (numeric and non‑numeric)
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Text");
            sheet.Cells["A4"].PutValue(25);
            sheet.Cells["A5"].PutValue(30);
            sheet.Cells["A6"].PutValue("Another");
            sheet.Cells["A7"].PutValue(45);

            // Enumerate column A, collect only numeric values, and write them sequentially to column B
            int targetRow = 2; // start writing numeric values from B2
            for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
            {
                Cell cell = sheet.Cells[row, 0]; // column A (index 0)
                if (cell.Type == CellValueType.IsNumeric)
                {
                    sheet.Cells[targetRow, 1].PutValue(cell.DoubleValue); // column B (index 1)
                    targetRow++;
                }
            }

            // Determine the range that contains the collected numeric values in column B
            int firstDataRow = 2;
            int lastDataRow = targetRow - 1;
            string numericRange = $"=Sheet1!$B${firstDataRow}:$B${lastDataRow}";

            // Add a column chart and set its series to the numeric range
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add(numericRange, true);

            // Optionally set category labels (using the original rows as categories)
            // Here we just use the row numbers as categories
            string categoryRange = $"=Sheet1!$A${firstDataRow}:$A${lastDataRow}";
            chart.NSeries.CategoryData = categoryRange;

            // Save the workbook
            workbook.Save("EnumeratedSeries.xlsx");
        }
    }
}