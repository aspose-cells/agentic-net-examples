using Aspose.Cells;
using Aspose.Cells.Charts;

class ApplyChartStyle
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the charts
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Add first chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart1 = sheet.Charts[chartIndex];
        chart1.NSeries.Add("B2:B4", false);
        chart1.NSeries.CategoryData = "A2:A4";

        // Add second chart
        chartIndex = sheet.Charts.Add(ChartType.Column, 5, 10, 20, 18);
        Chart chart2 = sheet.Charts[chartIndex];
        chart2.NSeries.Add("B2:B4", false);
        chart2.NSeries.CategoryData = "A2:A4";

        // Apply predefined style 'Style1' (style number 1) to all charts in the workbook
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart ch in ws.Charts)
            {
                ch.Style = 1; // Built‑in style index for Style1
            }
        }

        // Save the workbook
        workbook.Save("WorkbookWithStyledCharts.xlsx");
    }
}