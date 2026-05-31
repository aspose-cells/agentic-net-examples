using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample sales data (Category in column A, Values in column B)
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            int[] sales = { 120, 150, 130, 170, 160 };
            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 2, 0].PutValue(months[i]);   // Column A
                sheet.Cells[i + 2, 1].PutValue(sales[i]);   // Column B
            }

            // Create a named range called "SalesData" that refers to the sales values (B2:B6)
            int nameIndex = workbook.Worksheets.Names.Add("SalesData");
            // Use absolute reference to the range on the current sheet
            workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$B$2:$B$6";

            // Add a chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Link the chart series to the named range "SalesData"
            // The Add method accepts a named range directly
            chart.NSeries.Add("SalesData", true);

            // Optionally set the category (X‑axis) data to the month names
            chart.NSeries.CategoryData = $"={sheet.Name}!$A$2:$A$6";

            // Give the chart a title
            chart.Title.Text = "Monthly Sales";

            // Save the workbook
            workbook.Save("ChartWithNamedRange.xlsx");
        }
    }
}