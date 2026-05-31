using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Properties;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDynamicChartDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data (categories in column A, values in column B)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                string[] categories = { "A", "B", "C", "D", "E" };
                double[] values = { 10, 20, 30, 40, 50 };
                for (int i = 0; i < categories.Length; i++)
                {
                    sheet.Cells[i + 2, 0].PutValue(categories[i]); // A2:A6
                    sheet.Cells[i + 2, 1].PutValue(values[i]);    // B2:B6
                }

                // 3. Set a dynamic array formula in D1 that filters non‑empty values from column B
                //    The formula will spill into D1:D5 (or fewer rows if data changes)
                Cell dynamicCell = sheet.Cells["D1"];
                dynamicCell.SetDynamicArrayFormula("=FILTER(B2:B6, A2:A6<>\"\")", new FormulaParseOptions(), true);

                // 4. Create a named range that points to the spill range of the dynamic array (D1#)
                int nameIndex = sheet.Workbook.Worksheets.Names.Add("ChartData");
                Name chartDataName = sheet.Workbook.Worksheets.Names[nameIndex];
                chartDataName.RefersTo = "=Sheet1!$D$1#";

                // 5. Retrieve the range via the Name object (demonstrates GetRange rule)
                AsposeRange chartDataRange = chartDataName.GetRange();
                Console.WriteLine($"Named range 'ChartData' refers to address: {chartDataRange.RefersTo}");

                // 6. Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 0, 25, 8);
                Chart chart = sheet.Charts[chartIndex];

                // 7. Link the chart to the named range using SetChartDataRange
                //    The second parameter (true) indicates that the series are plotted by column
                chart.SetChartDataRange("ChartData", true);

                // 8. Set chart title and enable legend
                chart.Title.Text = "Dynamic Chart Linked to Named Range";
                chart.ShowLegend = true;

                // 9. Save the initial workbook
                string initialPath = "DynamicChart_Initial.xlsx";
                workbook.Save(initialPath);
                Console.WriteLine($"Saved: {Path.GetFullPath(initialPath)}");

                // ------------------------------------------------------------
                // 10. Update the underlying data programmatically (change B4 value)
                sheet.Cells["B4"].PutValue(99); // Change value for category "C"

                // 11. Refresh dynamic array formulas so the spill range updates
                workbook.RefreshDynamicArrayFormulas(true);

                // 12. Recalculate the chart to reflect the new data points
                chart.Calculate();

                // 13. Save the workbook after the update
                string updatedPath = "DynamicChart_Updated.xlsx";
                workbook.Save(updatedPath);
                Console.WriteLine($"Saved: {Path.GetFullPath(updatedPath)}");

                Console.WriteLine("Workbook saved. Dynamic chart reflects updated data.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}