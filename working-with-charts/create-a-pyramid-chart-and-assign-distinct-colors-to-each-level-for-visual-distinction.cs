// Title: C# Example: Create a Pyramid Chart with Unique Colors per Level Using Aspose.Cells
// Description: This Aspose.Cells for .NET sample builds a new workbook, fills A1:D4 with hierarchical data, adds a Pyramid chart, binds the series to B2:D4 and categories to A2:A4, then assigns solid red, green, and blue fills to each series before saving the file as PyramidChartDistinctColors.xlsx.
// Keywords: Aspose.Cells | C# pyramid chart | .NET Excel chart example | pyramid chart series colors | solid fill color Aspose.Cells | custom chart colors C# | Excel pyramid chart sample | chart customization Aspose | GitHub Aspose.Cells example | coding‑agent snippet
// Common Searches: Aspose.Cells create pyramid chart C# | how to set different colors for pyramid chart levels | customize series colors in Aspose.Cells chart | sample code for pyramid chart with distinct colors | Aspose.Cells .NET chart color fill example
// Developer Intent: Generate a pyramid chart in Excel and apply a distinct solid color to each level using Aspose.Cells for .NET.
// Use Cases: Display hierarchical sales or inventory data with a color‑coded pyramid for quick visual comparison. | Produce presentation‑ready Excel reports where each pyramid tier is highlighted with a unique color. | Automate monthly dashboards that include a pyramid chart with separate colors per series to improve readability.
// AI Prompts: Write C# code with Aspose.Cells that creates a three‑level pyramid chart and assigns red, green, and blue solid fills to each series. | Show how to replace solid fills with gradient fills for each level of a pyramid chart in Aspose.Cells. | Explain how to add data labels, a legend, and custom axis titles to the colored pyramid chart generated with Aspose.Cells.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This Aspose.Cells for .NET sample builds a new workbook, fills A1:D4 with hierarchical data, adds a Pyramid chart, binds the series to B2:D4 and categories to A2:A4, then assigns solid red, green, and blue fills to each series before saving the file as PyramidChartDistinctColors.xlsx.
class PyramidChartDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pyramid chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Level1");
        sheet.Cells["C1"].PutValue("Level2");
        sheet.Cells["D1"].PutValue("Level3");

        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(10);

        sheet.Cells["C2"].PutValue(20);
        sheet.Cells["C3"].PutValue(15);
        sheet.Cells["C4"].PutValue(5);

        sheet.Cells["D2"].PutValue(10);
        sheet.Cells["D3"].PutValue(5);
        sheet.Cells["D4"].PutValue(2);

        // Add a Pyramid chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Pyramid, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart series and categories
        chart.NSeries.Add("B2:D4", true);          // Values for all levels
        chart.NSeries.CategoryData = "A2:A4";      // Category labels

        // Assign distinct solid colors to each series (each level)
        Color[] levelColors = new Color[] { Color.Red, Color.Green, Color.Blue };
        for (int i = 0; i < chart.NSeries.Count && i < levelColors.Length; i++)
        {
            Series series = chart.NSeries[i];
            series.Area.FillFormat.FillType = FillType.Solid;
            series.Area.FillFormat.SolidFill.Color = levelColors[i];
        }

        // Save the workbook with the pyramid chart
        workbook.Save("PyramidChartDistinctColors.xlsx");
    }
}
