// Title: How to set a fixed major unit of 10 on the primary Y‑axis of a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Create an Excel workbook, add sample data, generate a column chart, and configure the primary Y‑axis major unit to 10 with Aspose.Cells in C#. | Modify an existing Aspose.Cells chart to turn off automatic major‑unit calculation and assign a constant major unit value of 10 to the value axis. | Produce a column chart where Y‑axis grid lines are spaced every 10 units and save the workbook as an .xlsx file using C#.
// Common Searches: Aspose.Cells C# set primary Y axis major unit interval to 10 | how to change Y axis grid spacing in a column chart with Aspose.Cells .NET | disable automatic major unit calculation on chart value axis Aspose.Cells | fixed major unit for chart Y axis example in C# | customize Y axis interval for Excel column chart using Aspose.Cells
// Tags: Aspose.Cells chart value axis major unit | C# fixed Y axis interval | column chart custom Y axis Aspose.Cells | Excel chart major unit configuration .NET | disable automatic major unit Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with sample data, adds a column chart, disables automatic major‑unit calculation on the primary Y‑axis, sets the major unit to 10, and saves the file as ChartWithCustomYAxisMajorUnit.xlsx.
class SetPrimaryYAxisMajorUnit
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the chart
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");
        for (int i = 2; i <= 6; i++)
        {
            cells[$"A{i}"].PutValue("Item " + (i - 1));
            cells[$"B{i}"].PutValue((i - 1) * 12); // sample values
        }

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B6", true);
        chart.NSeries.CategoryData = "A2:A6";

        // Access the primary Y axis (ValueAxis) and set the major unit to 10
        chart.ValueAxis.IsAutomaticMajorUnit = false; // disable automatic calculation
        chart.ValueAxis.MajorUnit = 10;               // set major unit interval

        // Save the workbook
        workbook.Save("ChartWithCustomYAxisMajorUnit.xlsx");
    }
}
