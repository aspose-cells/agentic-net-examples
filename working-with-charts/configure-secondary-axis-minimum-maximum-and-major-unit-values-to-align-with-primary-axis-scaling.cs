// Title: Set secondary axis min, max, and major unit to match primary axis in Aspose.Cells chart (C#)
// Description: Demonstrates how to create a workbook with a column chart, disable automatic scaling on the primary value axis, assign explicit MinValue, MaxValue, and MajorUnit, and then copy those settings to the secondary value axis so both axes share identical ranges before saving the file.
// Keywords: Aspose.Cells secondary axis scaling | C# chart axis manual range | ValueAxis IsAutomaticMinValue | SecondValueAxis MinValue MaxValue | Aspose.Cells chart axis alignment | set major unit Aspose.Cells | column chart multi‑axis Aspose | Aspose.Cells for .NET chart example
// Common Searches: How to set secondary axis minimum value in Aspose.Cells C# | Copy primary axis scaling to secondary axis Aspose.Cells | Aspose.Cells chart manual axis range example | Disable automatic axis scaling Aspose.Cells | Set major unit for secondary value axis Aspose.Cells
// Developer Intent: Programmatically configure the secondary value axis of an Aspose.Cells chart to use the same MinValue, MaxValue, and MajorUnit as the primary axis.
// Use Cases: Building a column chart with primary and secondary series that require identical scales for accurate visual comparison. | Generating financial or KPI reports where both axes must reflect the same range to avoid misinterpretation of data. | Automating chart formatting in a .NET application to ensure consistent axis settings across multiple workbooks.
// AI Prompts: Write C# code using Aspose.Cells to set the secondary value axis limits (min, max, major unit) identical to the primary axis for a column chart. | Explain the steps to disable automatic scaling and assign custom MinValue, MaxValue, and MajorUnit to both primary and secondary axes in an Aspose.Cells chart. | Provide an Aspose.Cells for .NET example that copies primary axis scaling properties to the secondary axis in a multi‑axis chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook with a column chart, disable automatic scaling on the primary value axis, assign explicit MinValue, MaxValue, and MajorUnit, and then copy those settings to the secondary value axis so both axes share identical ranges before saving the file.
public class SecondaryAxisScalingDemo
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for primary and secondary series
        cells["A1"].PutValue("Category");
        cells["A2"].PutValue("Jan");
        cells["A3"].PutValue("Feb");
        cells["A4"].PutValue("Mar");

        cells["B1"].PutValue("Primary");
        cells["B2"].PutValue(100);
        cells["B3"].PutValue(200);
        cells["B4"].PutValue(300);

        cells["C1"].PutValue("Secondary");
        cells["C2"].PutValue(400);
        cells["C3"].PutValue(500);
        cells["C4"].PutValue(600);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series: first for primary axis, second for secondary axis
        chart.NSeries.Add("B2:B4", true);   // Primary series
        chart.NSeries.Add("C2:C4", true);   // Secondary series
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary value axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Configure primary value axis scaling (min, max, major unit)
        Axis primaryAxis = chart.ValueAxis;
        primaryAxis.IsAutomaticMinValue = false;
        primaryAxis.IsAutomaticMaxValue = false;
        primaryAxis.IsAutomaticMajorUnit = false;
        primaryAxis.MinValue = 0;          // Minimum value
        primaryAxis.MaxValue = 800;        // Maximum value
        primaryAxis.MajorUnit = 200;       // Major unit interval

        // Align secondary axis scaling with the primary axis
        Axis secondaryAxis = chart.SecondValueAxis;
        secondaryAxis.IsAutomaticMinValue = false;
        secondaryAxis.IsAutomaticMaxValue = false;
        secondaryAxis.IsAutomaticMajorUnit = false;
        secondaryAxis.MinValue = primaryAxis.MinValue;
        secondaryAxis.MaxValue = primaryAxis.MaxValue;
        secondaryAxis.MajorUnit = primaryAxis.MajorUnit;

        // Save the workbook to a file
        string outputPath = "SecondaryAxisScalingDemo.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
