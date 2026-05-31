using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ComboChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate sample data ----------
            // Column A : Dates (categories)
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["A2"].PutValue(DateTime.Today.AddDays(-4).ToShortDateString());
            sheet.Cells["A3"].PutValue(DateTime.Today.AddDays(-3).ToShortDateString());
            sheet.Cells["A4"].PutValue(DateTime.Today.AddDays(-2).ToShortDateString());
            sheet.Cells["A5"].PutValue(DateTime.Today.AddDays(-1).ToShortDateString());
            sheet.Cells["A6"].PutValue(DateTime.Today.ToShortDateString());

            // Column B : Volume (to be shown as Area)
            sheet.Cells["B1"].PutValue("Volume");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);
            sheet.Cells["B5"].PutValue(170);
            sheet.Cells["B6"].PutValue(160);

            // Column C : Price (to be shown as Line)
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["C2"].PutValue(45.5);
            sheet.Cells["C3"].PutValue(46.2);
            sheet.Cells["C4"].PutValue(44.8);
            sheet.Cells["C5"].PutValue(47.1);
            sheet.Cells["C6"].PutValue(46.5);

            // ---------- Add a combo chart ----------
            // Start with an Area chart; later we will change one series to Line
            int chartIdx = sheet.Charts.Add(ChartType.Area, 8, 0, 25, 10);
            Chart chart = sheet.Charts[chartIdx];
            chart.Title.Text = "Volume (Area) & Price (Line)";

            // Set the category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A6";

            // Add the first series – Volume (Area)
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries[0].Name = "Volume";

            // Add the second series – Price (Line)
            chart.NSeries.Add("C2:C6", true);
            chart.NSeries[1].Name = "Price";

            // Change the second series type to Line to create a combo chart
            chart.NSeries[1].Type = ChartType.Line;

            // Save the workbook
            workbook.Save("ComboAreaLineChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}