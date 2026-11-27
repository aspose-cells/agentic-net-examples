// Title: Configure chart axes and freeze the header row with axis labels using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart category axis title and major unit | Freeze first row in worksheet after adding chart with Aspose.Cells .NET | Set value axis minimum and maximum in Aspose.Cells chart C# | Create column chart with custom axis limits and freeze panes using Aspose.Cells | Aspose.Cells example freezing header row containing chart labels
// Tags: configure chart axes Aspose.Cells C# | set chart axis major unit Aspose.Cells | freeze panes worksheet Aspose.Cells | column chart data binding Aspose.Cells | custom axis titles Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills it with monthly sales data, adds a column chart, customizes the category and value axes (titles, major/minor units, min/max values), freezes the first data row that holds the axis labels, and saves the file as ChartWithFrozenHeader.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and get the first worksheet.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "SalesData";

            // -----------------------------------------------------------------
            // 2. Populate sample data that will be used for the chart.
            // -----------------------------------------------------------------
            worksheet.Cells["A1"].PutValue("Month");
            worksheet.Cells["B1"].PutValue("Sales");

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
            double[] sales = { 1200, 1500, 1800, 1300, 1700, 1600 };

            for (int i = 0; i < months.Length; i++)
            {
                worksheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A – Month
                worksheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B – Sales
            }

            // -----------------------------------------------------------------
            // 3. Add a column chart and bind it to the data range.
            // -----------------------------------------------------------------
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B7", true);
            chart.NSeries.CategoryData = "A2:A7";

            // -----------------------------------------------------------------
            // 4. Configure chart titles and axes.
            // -----------------------------------------------------------------
            chart.Title.Text = "Monthly Sales";
            chart.Title.IsVisible = true;

            Axis categoryAxis = chart.CategoryAxis;
            categoryAxis.Title.Text = "Month";
            categoryAxis.Title.IsVisible = true;
            categoryAxis.HasMultiLevelLabels = false;
            categoryAxis.IsAutomaticMajorUnit = false;
            categoryAxis.MajorUnit = 1;
            categoryAxis.TickLabelPosition = TickLabelPositionType.NextToAxis;

            Axis valueAxis = chart.ValueAxis;
            valueAxis.Title.Text = "Sales ($)";
            valueAxis.Title.IsVisible = true;
            valueAxis.IsAutomaticMinValue = false;
            valueAxis.MinValue = 0;
            valueAxis.IsAutomaticMaxValue = false;
            valueAxis.MaxValue = 2000;
            valueAxis.IsAutomaticMajorUnit = false;
            valueAxis.MajorUnit = 500;
            valueAxis.IsAutomaticMinorUnit = false;
            valueAxis.MinorUnit = 100;
            valueAxis.DisplayUnit = DisplayUnitType.None;

            // -----------------------------------------------------------------
            // 5. Freeze the header row that contains the axis labels.
            // -----------------------------------------------------------------
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // Row = 1 (cell A2) freezes the first row (index 0).
            worksheet.FreezePanes(1, 0, 1, 0);

            // -----------------------------------------------------------------
            // 6. Save the workbook.
            // -----------------------------------------------------------------
            string outputPath = "ChartWithFrozenHeader.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
