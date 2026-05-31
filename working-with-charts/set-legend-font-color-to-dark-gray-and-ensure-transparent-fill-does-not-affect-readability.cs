using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetLegendFontColor
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Ensure the legend is displayed
        chart.ShowLegend = true;

        // Set legend entry font color to dark gray and make background transparent
        foreach (LegendEntry entry in chart.Legend.LegendEntries)
        {
            entry.Font.Color = Color.DarkGray;          // Dark gray font
            entry.BackgroundMode = BackgroundMode.Transparent; // Transparent background
            entry.IsTextNoFill = false;                // Ensure text fill is applied
        }

        // Save the workbook
        workbook.Save("LegendFontDarkGray.xlsx");
    }
}