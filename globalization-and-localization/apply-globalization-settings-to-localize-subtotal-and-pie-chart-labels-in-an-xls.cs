using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Settings;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // -------------------------------------------------
        // 1. Prepare data for Subtotal demonstration
        // -------------------------------------------------
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Amount");
        cells["A2"].PutValue("A");
        cells["B2"].PutValue(10);
        cells["A3"].PutValue("B");
        cells["B3"].PutValue(20);
        cells["A4"].PutValue("C");
        cells["B4"].PutValue(30);

        // Apply Subtotal (group by column 0, use Sum function)
        CellArea area = CellArea.CreateCellArea(0, 0, 4, 1);
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

        // -------------------------------------------------
        // 2. Globalize Subtotal label (e.g., change "Sum" to Chinese "合计")
        // -------------------------------------------------
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
        globalization.SetTotalName(ConsolidationFunction.Sum, "合计"); // custom total name for Sum

        // -------------------------------------------------
        // 3. Prepare data for Pie Chart demonstration
        // -------------------------------------------------
        // (Re‑use the same data range A2:B4)
        int chartIdx = sheet.Charts.Add(ChartType.Pie, 7, 0, 20, 10);
        Chart pieChart = sheet.Charts[chartIdx];
        pieChart.NSeries.Add("B2:B4", true);          // values
        pieChart.NSeries.CategoryData = "A2:A4";     // categories
        pieChart.Title.Text = "销售分布";               // chart title in Chinese

        // -------------------------------------------------
        // 4. Globalize Pie Chart labels (Series name, Legend Total, "Other" slice)
        // -------------------------------------------------
        SettableChartGlobalizationSettings chartGlob = new SettableChartGlobalizationSettings();
        chartGlob.SetSeriesName("系列");               // series name
        chartGlob.SetLegendTotalName("总计");          // legend total label
        chartGlob.SetOtherName("其他");               // label for "Other" slice

        // Attach chart globalization settings to the workbook settings
        globalization.ChartSettings = chartGlob;
        workbook.Settings.GlobalizationSettings = globalization;

        // -------------------------------------------------
        // 5. Save the workbook
        // -------------------------------------------------
        workbook.Save("GlobalizedSubtotalAndPieChart.xlsx");
    }
}