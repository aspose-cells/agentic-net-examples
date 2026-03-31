using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace CompositeChartDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (Category in column A, two data series in B and C)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // -----------------------------------------------------------------
            // First chart: Column chart for Series1
            // -----------------------------------------------------------------
            int chartIdx1 = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart1 = sheet.Charts[chartIdx1];
            chart1.NSeries.Add("B2:B4", true);               // Values
            chart1.NSeries.CategoryData = "A2:A4";           // Categories
            chart1.Title.Text = "Column Chart - Series1";

            // -----------------------------------------------------------------
            // Second chart: Line chart for Series2
            // -----------------------------------------------------------------
            int chartIdx2 = sheet.Charts.Add(ChartType.Line, 16, 0, 26, 5);
            Chart chart2 = sheet.Charts[chartIdx2];
            chart2.NSeries.Add("C2:C4", true);
            chart2.NSeries.CategoryData = "A2:A4";
            chart2.Title.Text = "Line Chart - Series2";

            // -----------------------------------------------------------------
            // Composite chart: combine both series into a single chart
            // -----------------------------------------------------------------
            int compIdx = sheet.Charts.Add(ChartType.Column, 27, 0, 37, 5);
            Chart composite = sheet.Charts[compIdx];
            composite.Title.Text = "Composite Chart";

            // Add Series1 (column)
            composite.NSeries.Add("B2:B4", true);
            composite.NSeries[0].Name = "Series1";

            // Add Series2 (line) – after adding, change its chart type
            composite.NSeries.Add("C2:C4", true);
            composite.NSeries[1].Name = "Series2";
            composite.NSeries[1].Type = ChartType.Line; // Convert second series to line

            // Optional: remove the original individual charts
            // Note: Removing by index must start from the highest index to avoid shifting
            sheet.Charts.RemoveAt(chartIdx2);
            sheet.Charts.RemoveAt(chartIdx1);

            // Save the workbook with the composite chart
            workbook.Save("CompositeChartOutput.xlsx", SaveFormat.Xlsx);
        }
    }
}