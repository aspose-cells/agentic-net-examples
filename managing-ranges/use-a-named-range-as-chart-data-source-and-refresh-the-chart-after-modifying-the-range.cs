using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsNamedRangeChartDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (Category in column A, Values in column B)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Create a named range that refers to the data area
            int nameIndex = workbook.Worksheets.Names.Add("MyData");
            // Note: RefersTo must start with '='
            workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$B$4";

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Use the named range as the chart data source
            // The second parameter (true) indicates that data is plotted by columns
            chart.NSeries.Add("MyData", true);

            // Optional: display the data range that the chart is using
            Console.WriteLine("Initial chart data range: " + chart.GetChartDataRange());

            // Save the workbook after initial creation
            workbook.Save("NamedRangeChart_Initial.xlsx");

            // ----- Modify the data in the named range -----
            // Change some values to simulate data update
            sheet.Cells["B2"].PutValue(15); // Update value for category A
            sheet.Cells["B3"].PutValue(25); // Update value for category B
            sheet.Cells["B4"].PutValue(35); // Update value for category C

            // Refresh the chart so it reflects the updated data
            // Calculate the chart to re-evaluate its data source
            chart.Calculate();

            // Verify that the chart detects data change
            bool dataChanged = chart.IsChartDataChanged();
            Console.WriteLine("Chart data changed after modification: " + dataChanged);

            // Save the workbook after data modification and chart refresh
            workbook.Save("NamedRangeChart_Updated.xlsx");
        }
    }
}