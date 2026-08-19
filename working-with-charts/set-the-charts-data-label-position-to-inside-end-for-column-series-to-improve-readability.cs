// Title: Aspose.Cells .NET: Set Column Chart Series Data Labels to Inside End Position (C#)
// Description: Shows how to create a workbook, add a column chart, enable data labels, and place them at the InsideEnd location for clearer visualization using Aspose.Cells for .NET (C#).
// Keywords: Aspose.Cells | C# | .NET | column chart | data label position | InsideEnd | LabelPositionType | chart series | Excel automation | chart customization
// Common Searches: How to set InsideEnd data label position in Aspose.Cells column chart | Aspose.Cells C# set chart data labels inside end | Change data label location for column series using Aspose.Cells | Aspose.Cells chart label position example C#
// Developer Intent: The developer wants to position the data labels of a column‑chart series at the InsideEnd location to improve readability.
// Use Cases: Generating Excel reports where column values are displayed inside the top of each bar. | Automating workbook creation with charts that show data labels at the InsideEnd position for quick visual analysis. | Customizing chart appearance in .NET applications that use Aspose.Cells to produce professional‑looking spreadsheets.
// AI Prompts: Provide C# code that sets a column chart series data label position to InsideEnd using Aspose.Cells. | Show an Aspose.Cells example that enables data labels and positions them inside the end of columns. | Explain how to customize chart data label positions, including InsideEnd, for different chart types in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLabelPosition
{
    // Shows how to create a workbook, add a column chart, enable data labels, and place them at the InsideEnd location for clearer visualization using Aspose.Cells for .NET (C#).
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Set data label position to InsideEnd for better readability
            chart.NSeries[0].DataLabels.Position = LabelPositionType.InsideEnd;

            // Save the workbook to a file
            workbook.Save("ColumnChart_With_InsideEndLabels.xlsx");
        }
    }
}
