// Title: Set Secondary Axis Min, Max, and Major Unit to Match Primary Axis in an Aspose.Cells Column Chart (C#)
// Description: Creates a workbook with category data and two series (small and large values), adds a column chart, plots the second series on the secondary value axis, disables automatic scaling, manually defines MinValue, MaxValue, and MajorUnit on the primary axis, copies those settings to the secondary axis, adds axis titles, and saves the file as ConfigureSecondaryAxis.xlsx.
// Keywords: Aspose.Cells | C# chart axis scaling | secondary value axis | primary axis min max | manual axis limits | column chart Aspose.Cells | ConfigureSecondaryAxis | Excel report generation | Aspose.Cells for .NET | chart axis major unit
// Common Searches: Aspose.Cells set secondary axis minimum value C# | copy primary axis scaling to secondary axis Aspose.Cells | how to set major unit for secondary axis in Aspose.Cells chart | disable automatic axis scaling Aspose.Cells C# | match secondary axis range with primary axis Aspose.Cells
// Developer Intent: Programmatically align the secondary value axis limits and tick interval with those of the primary axis in a column chart using Aspose.Cells for .NET.
// Use Cases: Combine small‑scale and large‑scale data series in one chart while keeping both axes on the same scale for easy visual comparison. | Generate automated Excel reports where consistent tick intervals across primary and secondary axes are required for regulatory or presentation standards.
// AI Prompts: Write C# code with Aspose.Cells that sets the secondary axis MinValue, MaxValue, and MajorUnit to the same values as the primary axis. | Explain step‑by‑step how to turn off automatic scaling and apply manual axis limits for both primary and secondary axes in an Aspose.Cells chart. | Provide a concise guide to synchronize secondary axis scaling with the primary axis in a column chart using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook with category data and two series (small and large values), adds a column chart, plots the second series on the secondary value axis, disables automatic scaling, manually defines MinValue, MaxValue, and MajorUnit on the primary axis, copies those settings to the secondary axis, adds axis titles, and saves the file as ConfigureSecondaryAxis.xlsx.
class ConfigureSecondaryAxis
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data
        cells["A1"].PutValue("Category");
        cells["A2"].PutValue("A");
        cells["A3"].PutValue("B");
        cells["A4"].PutValue("C");

        // Primary series values (small numbers)
        cells["B1"].PutValue("Primary Series");
        cells["B2"].PutValue(10);
        cells["B3"].PutValue(20);
        cells["B4"].PutValue(30);

        // Secondary series values (larger numbers)
        cells["C1"].PutValue("Secondary Series");
        cells["C2"].PutValue(1000);
        cells["C3"].PutValue(2000);
        cells["C4"].PutValue(3000);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series: first uses primary axis, second uses secondary axis
        chart.NSeries.Add("B2:B4", true);          // Primary series
        chart.NSeries.Add("C2:C4", true);          // Secondary series
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary value axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // ----- Configure primary value axis -----
        Axis primaryAxis = chart.ValueAxis;
        primaryAxis.IsAutomaticMinValue = false;
        primaryAxis.IsAutomaticMaxValue = false;
        primaryAxis.IsAutomaticMajorUnit = false;

        primaryAxis.MinValue = 0;      // Minimum
        primaryAxis.MaxValue = 40;     // Maximum
        primaryAxis.MajorUnit = 10;    // Major unit

        // ----- Configure secondary value axis to match primary scaling -----
        Axis secondaryAxis = chart.SecondValueAxis;
        secondaryAxis.IsAutomaticMinValue = false;
        secondaryAxis.IsAutomaticMaxValue = false;
        secondaryAxis.IsAutomaticMajorUnit = false;

        // Align secondary axis scaling with primary axis
        secondaryAxis.MinValue = primaryAxis.MinValue;
        secondaryAxis.MaxValue = primaryAxis.MaxValue;
        secondaryAxis.MajorUnit = primaryAxis.MajorUnit;

        // Optional: give titles to axes for clarity
        primaryAxis.Title.Text = "Primary Axis";
        secondaryAxis.Title.Text = "Secondary Axis";

        // Save the workbook
        workbook.Save("ConfigureSecondaryAxis.xlsx");
    }
}
