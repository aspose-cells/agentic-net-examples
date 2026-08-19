// Title: Add a secondary Value Y‑axis and bind a series in Aspose.Cells for .NET (C#)
// Description: This example builds a workbook, inserts month, sales and profit rows, creates a column chart, defines two data series and moves the profit series to a second value axis. The secondary axis is given a custom title and explicit min, max and major‑unit settings before the file is saved.
// Keywords: Aspose.Cells C# secondary axis | dual‑axis chart .NET | plot series on second value axis | customize chart axis range Aspose | column chart secondary Y axis | set axis title Aspose.Cells | secondary value axis configuration | Aspose.Cells chart example | C# Excel chart secondary axis
// Common Searches: how to create a secondary value axis in Aspose.Cells C# | assign a series to the second Y axis using Aspose.Cells | set minimum and maximum for a secondary chart axis .NET | dual axis column chart Aspose.Cells tutorial | custom axis title for secondary axis Aspose
// Developer Intent: Generate a column chart where one series is displayed on a second value axis with custom scaling and labeling.
// Use Cases: Compare revenue and profit in a single visual while keeping distinct scales. | Produce financial dashboards that require separate units for sales and margin. | Export Excel reports with dual‑axis charts for business intelligence tools.
// AI Prompts: Show C# code that adds a second value axis to an Aspose.Cells chart and assigns a series to it. | Explain how to set the title, minimum, maximum and major unit of a secondary Y‑axis in Aspose.Cells for .NET. | Provide steps to create a dual‑axis column chart with different data series on primary and secondary axes using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// This example builds a workbook, inserts month, sales and profit rows, creates a column chart, defines two data series and moves the profit series to a second value axis. The secondary axis is given a custom title and explicit min, max and major‑unit settings before the file is saved.
class SecondaryYAxisDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Month");
        worksheet.Cells["A2"].PutValue("Jan");
        worksheet.Cells["A3"].PutValue("Feb");
        worksheet.Cells["A4"].PutValue("Mar");

        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["B4"].PutValue(180);

        worksheet.Cells["C1"].PutValue("Profit");
        worksheet.Cells["C2"].PutValue(30);
        worksheet.Cells["C3"].PutValue(45);
        worksheet.Cells["C4"].PutValue(60);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Add two series: Sales (primary Y‑axis) and Profit (secondary Y‑axis)
        chart.NSeries.Add("B2:B4", true); // Series 0 – Sales
        chart.NSeries.Add("C2:C4", true); // Series 1 – Profit
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary Y‑axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Customize the secondary Y‑axis (value axis)
        Axis secondaryValueAxis = chart.SecondValueAxis;
        secondaryValueAxis.Title.Text = "Profit";
        secondaryValueAxis.MinValue = 0;
        secondaryValueAxis.MaxValue = 100;
        secondaryValueAxis.MajorUnit = 20;

        // Save the workbook
        workbook.Save("SecondaryYAxisDemo.xlsx");
    }
}
