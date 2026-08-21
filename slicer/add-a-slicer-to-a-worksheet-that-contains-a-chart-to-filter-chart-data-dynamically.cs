// Title: Add a slicer to a pivot‑based chart in C# using Aspose.Cells for .NET
// Description: Creates a workbook, fills it with fruit‑sales data, builds a pivot table, generates a column chart linked to the pivot, inserts a slicer for the "Fruit" field, and saves the file as ChartWithSlicer.xlsx. The slicer updates the chart automatically when selections change.
// Keywords: Aspose.Cells slicer example | C# pivot chart slicer | dynamic chart filtering .NET | add slicer to Excel chart programmatically | Aspose.Cells pivot table chart | Excel slicer API C# | interactive dashboard Aspose.Cells
// Common Searches: how to add a slicer to a pivot chart with Aspose.Cells | C# code for slicer linked to pivot table | Aspose.Cells example chart with slicer | filter chart data using slicer in .NET | Aspose.Cells dynamic dashboard tutorial
// Developer Intent: Programmatically attach a slicer to a pivot table so that a linked chart refreshes when the slicer selection changes.
// Use Cases: Create an interactive sales dashboard where users filter fruit categories via a slicer and see the chart update instantly. | Generate Excel reports that combine pivot tables, charts, and slicers for on‑the‑fly data exploration. | Build a reusable utility that adds a pivot table, chart, and associated slicers to any worksheet for dynamic analysis.
// AI Prompts: Generate C# code with Aspose.Cells to add a slicer for the "Year" field to an existing pivot chart and apply a custom style. | Explain how to force a chart to refresh after a slicer selection changes in Aspose.Cells for .NET. | Write a method that creates a pivot table, a column chart, and multiple slicers for different fields in a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Slicers;

// Creates a workbook, fills it with fruit‑sales data, builds a pivot table, generates a column chart linked to the pivot, inserts a slicer for the "Fruit" field, and saves the file as ChartWithSlicer.xlsx. The slicer updates the chart automatically when selections change.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data
        cells["A1"].PutValue("Fruit");
        cells["B1"].PutValue("Year");
        cells["C1"].PutValue("Amount");

        string[] fruits = { "Apple", "Banana", "Apple", "Banana", "Apple", "Banana" };
        int[] years   = { 2020,   2020,    2021,   2021,    2022,   2022 };
        int[] amounts = { 50,    70,      60,     80,      55,     85 };

        for (int i = 0; i < fruits.Length; i++)
        {
            cells[i + 1, 0].PutValue(fruits[i]);
            cells[i + 1, 1].PutValue(years[i]);
            cells[i + 1, 2].PutValue(amounts[i]);
        }

        // Add a pivot table based on the data
        int pivotIndex = sheet.PivotTables.Add("A1:C7", "E2", "Pivot1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Column, "Year");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a chart that uses the pivot table as its source
        int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 10);
        Chart chart = sheet.Charts[chartIndex];
        // The chart will automatically pick up the pivot table data
        chart.NSeries.Add(pivot.Name + "!Data", true);
        chart.Title.Text = "Sales by Fruit and Year";

        // Add a slicer linked to the pivot table for the "Fruit" field
        SlicerCollection slicers = sheet.Slicers;
        int slicerIndex = slicers.Add(pivot, "E12", "Fruit");
        Slicer slicer = slicers[slicerIndex];
        slicer.Caption = "Fruit Filter";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

        // Save the workbook
        workbook.Save("ChartWithSlicer.xlsx", SaveFormat.Xlsx);
    }
}
