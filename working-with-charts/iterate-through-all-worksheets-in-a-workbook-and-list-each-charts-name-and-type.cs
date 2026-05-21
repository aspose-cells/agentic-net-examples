using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartInfo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Add sample data and two charts to the first worksheet for demonstration
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Category");
            sheet1.Cells["A2"].PutValue("A");
            sheet1.Cells["A3"].PutValue("B");
            sheet1.Cells["A4"].PutValue("C");
            sheet1.Cells["B1"].PutValue("Value");
            sheet1.Cells["B2"].PutValue(10);
            sheet1.Cells["B3"].PutValue(20);
            sheet1.Cells["B4"].PutValue(30);

            // First chart (Column)
            int chartIdx1 = sheet1.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart1 = sheet1.Charts[chartIdx1];
            chart1.NSeries.Add("B2:B4", true);
            chart1.NSeries.CategoryData = "A2:A4";
            chart1.Name = "SalesColumn";

            // Second chart (Pie)
            int chartIdx2 = sheet1.Charts.Add(ChartType.Pie, 20, 0, 30, 5);
            Chart chart2 = sheet1.Charts[chartIdx2];
            chart2.NSeries.Add("B2:B4", true);
            chart2.NSeries.CategoryData = "A2:A4";
            chart2.Name = "SalesPie";

            // Add a second worksheet without charts
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            sheet2.Cells["A1"].PutValue("No charts here");

            // Iterate through all worksheets and list each chart's name and type
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // If the worksheet contains charts, iterate them
                foreach (Chart ch in ws.Charts)
                {
                    // Chart.Name may be null if not set; handle gracefully
                    string chartName = string.IsNullOrEmpty(ch.Name) ? "(unnamed)" : ch.Name;
                    Console.WriteLine($"Worksheet: {ws.Name}, Chart Name: {chartName}, Chart Type: {ch.Type}");
                }
            }

            // Save the workbook (save rule)
            workbook.Save("ChartInfoOutput.xlsx");
        }
    }
}