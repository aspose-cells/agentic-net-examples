// Title: Set the plot area background color of an Excel chart to light yellow using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to change the plot area fill color of a column chart to light yellow | Aspose.Cells .NET set solid background for chart plot area | example code for setting chart plot area background color in Excel with Aspose.Cells | apply light yellow fill to Excel chart plot area using Aspose.Cells C# | Aspose.Cells chart formatting plot area background color tutorial
// Tags: Aspose.Cells chart plot area fill format | set solid fill color for Excel chart area C# | light yellow plot area background Aspose.Cells | chart plot area border styling Aspose.Cells .NET | Excel chart formatting with Aspose.Cells C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a column chart, and uses the PlotArea.Area.FillFormat API to apply a solid light‑yellow fill (and optionally a black border) to the chart's plot area before saving the file as an XLSX document.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data for the chart
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");
        cells["A2"].PutValue("A");
        cells["A3"].PutValue("B");
        cells["A4"].PutValue("C");
        cells["B2"].PutValue(10);
        cells["B3"].PutValue(20);
        cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // ----- Set the plot area background to light yellow -----
        // Access the PlotArea -> Area -> FillFormat
        chart.PlotArea.Area.FillFormat.FillType = FillType.Solid;
        chart.PlotArea.Area.FillFormat.SolidFill.Color = Color.LightYellow;

        // (Optional) Add a thin black border around the plot area
        chart.PlotArea.Border.Color = Color.Black;

        // Save the workbook
        string outputFile = "ChartWithLightYellowPlotArea.xlsx";
        workbook.Save(outputFile, SaveFormat.Xlsx);
        Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
    }
}
