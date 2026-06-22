using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsNamedRangeChart
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A – Category labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Column B – Sales values (this range will be named)
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);
            sheet.Cells["B5"].PutValue(170);

            // Define a named range that refers to the sales data (excluding the header)
            int nameIndex = workbook.Worksheets.Names.Add("SalesData");
            // RefersTo must be a formula string starting with '='
            workbook.Worksheets.Names[nameIndex].RefersTo = $"=Sheet1!$B$2:$B$5";

            // (Optional) Define a named range for the category labels
            int catNameIndex = workbook.Worksheets.Names.Add("CategoryLabels");
            workbook.Worksheets.Names[catNameIndex].RefersTo = $"=Sheet1!$A$2:$A$5";

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Use the named range as the data source for the series.
            // The second parameter 'true' indicates that the data is plotted column‑by‑column.
            chart.NSeries.Add("SalesData", true);

            // Set the category (X‑axis) data using the named range for labels
            chart.NSeries.CategoryData = "CategoryLabels";

            // Give the chart a title
            chart.Title.Text = "Monthly Sales (Named Range)";

            // Save the workbook – the chart will automatically reflect any changes made to the named range cells
            workbook.Save("NamedRangeChart.xlsx");
        }
    }
}