using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class VerifyChartLocalization
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // ----- Localized source data -----
        // Japanese strings for categories, values, and labels
        worksheet.Cells["A1"].PutValue("カテゴリ");          // Header for categories
        worksheet.Cells["A2"].PutValue("商品A");            // Category 1
        worksheet.Cells["A3"].PutValue("商品B");            // Category 2
        worksheet.Cells["B1"].PutValue("売上");            // Header for values
        worksheet.Cells["B2"].PutValue(120);               // Value 1
        worksheet.Cells["B3"].PutValue(150);               // Value 2

        // Cells that contain the localized chart title and axis titles
        worksheet.Cells["C1"].PutValue("売上チャート");      // Chart title
        worksheet.Cells["D1"].PutValue("商品");            // Category (X) axis title
        worksheet.Cells["E1"].PutValue("金額 (千円)");      // Value (Y) axis title

        // ----- Create a chart -----
        // Add a column chart positioned from row 5, column 0 to row 20, column 10
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Bind data series and categories
        chart.NSeries.Add("B2:B3", true);          // Values
        chart.NSeries.CategoryData = "A2:A3";      // Categories

        // ----- Apply localized titles from source cells -----
        chart.Title.Text = worksheet.Cells["C1"].StringValue;               // Chart title
        chart.CategoryAxis.Title.Text = worksheet.Cells["D1"].StringValue; // X‑axis title
        chart.ValueAxis.Title.Text = worksheet.Cells["E1"].StringValue;    // Y‑axis title

        // ----- Verification output -----
        Console.WriteLine("Chart Title: " + chart.Title.Text);
        Console.WriteLine("Category Axis Title: " + chart.CategoryAxis.Title.Text);
        Console.WriteLine("Value Axis Title: " + chart.ValueAxis.Title.Text);

        // Save the workbook (lifecycle rule)
        workbook.Save("LocalizedChart.xlsx");
    }
}