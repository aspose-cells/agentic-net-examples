using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSeriesFromColumn
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column A with mixed data (numeric and non‑numeric)
            sheet.Cells["A1"].PutValue("Label");
            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["A3"].PutValue(10);          // numeric
            sheet.Cells["A4"].PutValue("Item 2");
            sheet.Cells["A5"].PutValue(20.5);        // numeric
            sheet.Cells["A6"].PutValue("Item 3");
            sheet.Cells["A7"].PutValue(30);          // numeric
            sheet.Cells["A8"].PutValue("Text");

            // Enumerate column A, collect only numeric values
            List<double> numericValues = new List<double>();
            int maxRow = sheet.Cells.MaxDataRow; // last row with data
            for (int row = 0; row <= maxRow; row++)
            {
                Cell cell = sheet.Cells[row, 0]; // column A (index 0)
                if (cell.Type == CellValueType.IsNumeric)
                {
                    numericValues.Add(cell.DoubleValue);
                }
            }

            // Write the collected numeric values to column C starting at C2
            int startRow = 1; // zero‑based index for row 2
            for (int i = 0; i < numericValues.Count; i++)
            {
                sheet.Cells[startRow + i, 2].PutValue(numericValues[i]); // column C (index 2)
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Build the address of the numeric range in column C (e.g., C2:C4)
            string numericRange = $"C{startRow + 1}:C{startRow + numericValues.Count}";
            // Add the series using the numeric range; true = vertical orientation
            chart.NSeries.Add(numericRange, true);

            // (Optional) Set category labels from column B if desired
            // sheet.Cells["B1"].PutValue("Category");
            // sheet.Cells["B2"].PutValue("A");
            // sheet.Cells["B3"].PutValue("B");
            // sheet.Cells["B4"].PutValue("C");
            // chart.NSeries.CategoryData = "B2:B4";

            // Save the workbook
            workbook.Save("ChartFromNumericColumn.xlsx");
        }
    }
}