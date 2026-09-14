// Title: How to add a slicer linked to a pivot table for dynamic chart filtering using Aspose.Cells in C#
// AI Prompts: Generate C# code with Aspose.Cells that creates sample data, a pivot table, a column chart, and inserts a slicer linked to the pivot table to filter the chart. | Show how to position a slicer at a specific cell, set its caption, and apply a light style using the Aspose.Cells API. | Provide a complete example that saves the workbook as an XLSX file after adding the slicer and chart.
// Common Searches: aspnet add slicer to Excel workbook using Aspose.Cells C# example | link slicer to pivot table to control chart data with Aspose.Cells .NET | set slicer caption and style programmatically in Aspose.Cells C# | create pivot table and column chart then add slicer for dynamic filtering Aspose.Cells | save workbook with slicer and chart as XLSX using Aspose.Cells C#
// Tags: Aspose.Cells add slicer to pivot table | Aspose.Cells link slicer to chart | C# create column chart from range Aspose.Cells | Aspose.Cells set slicer caption and style | Aspose.Cells save workbook as XLSX

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Slicers;

// Demonstrates using Aspose.Cells for .NET (C#) to create sample data, build a pivot table, generate a column chart, add a slicer linked to the pivot table, set its caption and style, and save the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the chart and pivot table
            cells["A1"].PutValue("Fruit");
            cells["B1"].PutValue("Sales");

            string[] fruits = { "Apple", "Banana", "Apple", "Banana", "Cherry" };
            int[] sales = { 120, 150, 130, 170, 200 };

            for (int i = 0; i < fruits.Length; i++)
            {
                cells[i + 2, 0].PutValue(fruits[i]);   // Column A
                cells[i + 2, 1].PutValue(sales[i]);   // Column B
            }

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

            // Refresh pivot cache and calculate data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a chart that uses the source data range (A2:B6) as its data source
            int chartIndex = sheet.Charts.Add(ChartType.Column, 0, 3, 20, 15);
            Chart chart = sheet.Charts[chartIndex];
            string dataRangeFormula = $"'{sheet.Name}'!$A$2:$B$6";
            chart.NSeries.Add(dataRangeFormula, true);
            chart.Title.Text = "Sales by Fruit";

            // Add a slicer linked to the pivot table to filter the chart dynamically
            SlicerCollection slicers = sheet.Slicers;
            int slicerIndex = slicers.Add(pivot, "F1", "Fruit"); // Place slicer at cell F1
            Slicer slicer = slicers[slicerIndex];
            slicer.Caption = "Fruit Filter";
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

            // Save the workbook
            workbook.Save("SlicerChartDemo.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
