using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class UpdateComboChartSeries
{
    static void Main()
    {
        try
        {
            // ---------- Create a new workbook and add sample data ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["C1"].PutValue("Profit");

            // Initial data rows
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["C2"].PutValue(30);

            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["C3"].PutValue(45);

            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["C4"].PutValue(60);

            // ---------- Add a Combo chart (Column + Line) ----------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // First series – Column (Sales)
            // Data area format: "SheetName!XValuesRange,YValuesRange"
            chart.NSeries.Add("Sheet1!$A$2:$A$4,Sheet1!$B$2:$B$4", true);
            // Second series – Line (Profit)
            chart.NSeries.Add("Sheet1!$A$2:$A$4,Sheet1!$C$2:$C$4", true);
            // Change the second series type to Line
            chart.NSeries[1].Type = ChartType.Line;

            // Save the workbook with the initial chart
            workbook.Save("ComboChart_Initial.xlsx");

            // ---------- Modify the underlying data (add a new row) ----------
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B5"].PutValue(210);
            sheet.Cells["C5"].PutValue(75);

            // ---------- Update the chart series to reflect the new data ----------
            chart.NSeries[0].XValues = "Sheet1!$A$2:$A$5";
            chart.NSeries[0].Values = "Sheet1!$B$2:$B$5";

            chart.NSeries[1].XValues = "Sheet1!$A$2:$A$5";
            chart.NSeries[1].Values = "Sheet1!$C$2:$C$5";

            // Recalculate the chart so the changes take effect
            chart.Calculate();

            // Save the workbook with the updated chart
            workbook.Save("ComboChart_Updated.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}