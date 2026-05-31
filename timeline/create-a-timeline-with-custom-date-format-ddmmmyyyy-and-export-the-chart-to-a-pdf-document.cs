using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data: Fruit, Date, Amount
        cells["A1"].PutValue("Fruit");
        cells["B1"].PutValue("Date");
        cells["C1"].PutValue("Amount");

        string[] fruits = { "Apple", "Banana", "Cherry", "Date" };
        DateTime[] dates = {
            new DateTime(2021, 1, 5),
            new DateTime(2021, 2, 10),
            new DateTime(2021, 3, 15),
            new DateTime(2021, 4, 20)
        };
        int[] amounts = { 100, 150, 200, 250 };

        for (int i = 0; i < fruits.Length; i++)
        {
            cells[i + 1, 0].PutValue(fruits[i]);   // Fruit column
            cells[i + 1, 1].PutValue(dates[i]);    // Date column
            cells[i + 1, 2].PutValue(amounts[i]);  // Amount column
        }

        // Create a date style with custom format dd-MMM-yyyy
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "dd-MMM-yyyy";

        // Apply the custom date style to the date cells
        for (int i = 1; i <= fruits.Length; i++)
        {
            cells[i, 1].SetStyle(dateStyle);
        }

        // Add a PivotTable based on the data range
        PivotTableCollection pivots = sheet.PivotTables;
        int pivotIndex = pivots.Add("A1:C5", "E1", "PivotTable1");
        PivotTable pivot = pivots[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Column, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a Timeline linked to the PivotTable using the Date field
        TimelineCollection timelines = sheet.Timelines;
        int timelineIndex = timelines.Add(pivot, "G1", "Date");
        Timeline timeline = timelines[timelineIndex];
        timeline.Caption = "Sales Timeline";

        // Add a line chart that uses the same data
        int chartIndex = sheet.Charts.Add(ChartType.Line, 15, 0, 30, 15);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("C2:C5", true);          // Values
        chart.NSeries.CategoryData = "B2:B5";      // Dates as categories

        // Set the X‑axis values format to match the custom date format
        chart.NSeries[0].XValuesFormatCode = "dd-MMM-yyyy";

        // Export the chart to a PDF file
        chart.ToPdf("TimelineChart.pdf");

        // Save the workbook (optional, for verification)
        workbook.Save("TimelineDemo.xlsx");
    }
}