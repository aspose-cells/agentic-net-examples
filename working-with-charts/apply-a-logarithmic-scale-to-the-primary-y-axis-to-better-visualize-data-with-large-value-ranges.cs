using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLogarithmicAxisDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with a wide range of values
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("Item 1");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Item 2");
            worksheet.Cells["B3"].PutValue(100);
            worksheet.Cells["A4"].PutValue("Item 3");
            worksheet.Cells["B4"].PutValue(1000);
            worksheet.Cells["A5"].PutValue("Item 4");
            worksheet.Cells["B5"].PutValue(10000);
            worksheet.Cells["A6"].PutValue("Item 5");
            worksheet.Cells["B6"].PutValue(100000);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Configure the primary Y axis (ValueAxis) to use a logarithmic scale
            Axis valueAxis = chart.ValueAxis;
            valueAxis.IsLogarithmic = true;      // Enable logarithmic scaling
            valueAxis.LogBase = 10;              // Set the logarithmic base (default is 10)
            valueAxis.MinValue = 1;              // Define minimum value for the axis
            valueAxis.MaxValue = 100000;         // Define maximum value for the axis

            // Optional: give the axis a title for clarity
            valueAxis.Title.Text = "Logarithmic Value Axis";

            // Save the workbook to an XLSX file
            workbook.Save("LogarithmicAxisDemo.xlsx");
        }
    }
}