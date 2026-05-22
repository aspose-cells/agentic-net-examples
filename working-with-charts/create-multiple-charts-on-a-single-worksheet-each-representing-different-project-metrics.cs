using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class MultipleChartsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "ProjectMetrics";

        // -------------------------------------------------
        // Populate sample data for three project metrics:
        // Revenue, Expenses and Profit for six months
        // -------------------------------------------------
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Revenue");
        sheet.Cells["C1"].PutValue("Expenses");
        sheet.Cells["D1"].PutValue("Profit");

        string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
        double[] revenue = { 12000, 15000, 13000, 16000, 17000, 18000 };
        double[] expenses = { 8000, 9000, 8500, 9500, 10000, 11000 };
        double[] profit = { 4000, 6000, 4500, 6500, 7000, 7000 };

        for (int i = 0; i < months.Length; i++)
        {
            int row = i + 2; // Data starts from row 2
            sheet.Cells[row, 0].PutValue(months[i]);      // Column A
            sheet.Cells[row, 1].PutValue(revenue[i]);     // Column B
            sheet.Cells[row, 2].PutValue(expenses[i]);    // Column C
            sheet.Cells[row, 3].PutValue(profit[i]);      // Column D
        }

        // -------------------------------------------------
        // Chart 1: Column chart showing Revenue by month
        // -------------------------------------------------
        int chartIndex1 = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
        Chart revenueChart = sheet.Charts[chartIndex1];
        revenueChart.Title.Text = "Revenue by Month";
        revenueChart.NSeries.Add("B2:B7", true);          // Values
        revenueChart.NSeries.CategoryData = "A2:A7";     // Categories

        // -------------------------------------------------
        // Chart 2: Line chart showing Expenses by month
        // -------------------------------------------------
        int chartIndex2 = sheet.Charts.Add(ChartType.Line, 5, 6, 20, 11);
        Chart expensesChart = sheet.Charts[chartIndex2];
        expensesChart.Title.Text = "Expenses by Month";
        expensesChart.NSeries.Add("C2:C7", true);
        expensesChart.NSeries.CategoryData = "A2:A7";

        // -------------------------------------------------
        // Chart 3: Pie chart showing Profit distribution
        // -------------------------------------------------
        int chartIndex3 = sheet.Charts.Add(ChartType.Pie, 22, 0, 35, 5);
        Chart profitChart = sheet.Charts[chartIndex3];
        profitChart.Title.Text = "Profit Distribution";
        profitChart.NSeries.Add("D2:D7", true);
        profitChart.NSeries.CategoryData = "A2:A7";

        // -------------------------------------------------
        // Save the workbook with all charts embedded
        // -------------------------------------------------
        workbook.Save("ProjectMetrics.xlsx", SaveFormat.Xlsx);
    }
}