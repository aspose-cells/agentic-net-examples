using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsNamedRangeChartDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (A1:B4)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Create a named range that refers to the data area
            int nameIndex = workbook.Worksheets.Names.Add("MyData");
            // Note: RefersTo must start with '='
            workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$A$1:$B$4";

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Use the named range as the chart data source
            // The second parameter (true) indicates that data is plotted by column
            chart.NSeries.Add("MyData", true);

            // Save the workbook after initial creation
            workbook.Save("NamedRangeChart_Initial.xlsx");

            // ----- Modify the data in the named range -----
            // Change the value of cell B2 (second row, second column of the range)
            sheet.Cells["B2"].PutValue(50);

            // Refresh the chart to reflect the updated data
            // Calculate the chart; this updates internal data caches
            chart.Calculate();

            // Optional: verify that the chart detects data change
            bool changed = chart.IsChartDataChanged();
            Console.WriteLine("Chart data changed after modification: " + changed);

            // Save the workbook after modification and refresh
            workbook.Save("NamedRangeChart_Updated.xlsx");
        }
    }
}