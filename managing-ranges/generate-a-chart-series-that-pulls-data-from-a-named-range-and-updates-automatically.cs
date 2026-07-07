using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsNamedRangeChart
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Define a named range that refers to the sales values (B2:B4)
            int nameIndex = workbook.Worksheets.Names.Add("SalesData");
            // The RefersTo string must start with '=' and include the sheet name
            workbook.Worksheets.Names[nameIndex].RefersTo = $"=Sheet1!$B$2:$B$4";

            // Define a named range for the category axis (A2:A4)
            int catIndex = workbook.Worksheets.Names.Add("MonthCategories");
            workbook.Worksheets.Names[catIndex].RefersTo = $"=Sheet1!$A$2:$A$4";

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Use the named range for the series values
            // The Add method accepts the name directly; the chart will reference the range dynamically
            chart.NSeries.Add("SalesData", true);

            // Set the category (X‑axis) data using the named range
            chart.NSeries.CategoryData = "MonthCategories";

            // Optional: give the series a friendly name (also can be a named range)
            chart.NSeries[0].Name = "Quarter 1 Sales";

            // Set a chart title
            chart.Title.Text = "Quarterly Sales";

            // Save the workbook – the chart will update automatically when the underlying cells change
            workbook.Save("NamedRangeChart.xlsx");
        }
    }
}