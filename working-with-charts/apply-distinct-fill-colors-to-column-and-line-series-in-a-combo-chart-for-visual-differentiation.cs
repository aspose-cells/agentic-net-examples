// Title: Aspose.Cells .NET: Apply Different Fill Colors to Column and Line Series in a Combo Chart (C#)
// Description: Creates a workbook with sales and profit data, adds a combo chart, converts the second series to a line type, assigns a blue solid fill to the column series and a red stroke to the line series, then saves the file as an Excel workbook.
// Keywords: Aspose.Cells combo chart color C# | set column series fill Aspose.Cells | line series color Aspose.Cells .NET | custom chart series colors Excel | Aspose.Cells chart styling US | Aspose.Cells chart styling India
// Common Searches: how to set column series fill color in Aspose.Cells combo chart | change line series color in Aspose.Cells chart C# | Aspose.Cells apply different colors to combo chart series | C# code for distinct colors in Excel combo chart using Aspose.Cells | Aspose.Cells chart formatting tutorial
// Developer Intent: Assign separate colors to the column and line series of a combo chart.
// Use Cases: Generate a sales‑vs‑profit report where columns are blue and the profit line is red for quick visual comparison. | Automate Excel workbook creation that highlights each data series with a unique RGB fill or stroke. | Produce presentation‑ready charts that follow corporate color guidelines by programmatically styling series.
// AI Prompts: Show C# code to give a column series a solid blue fill and a line series a red line color in an Aspose.Cells combo chart. | Provide an Aspose.Cells example that changes the second series to a line type and sets its SeriesLines.Color. | Explain how to customize fill and stroke colors for different series in a combo chart using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsComboChartDemo
{
    // Creates a workbook with sales and profit data, adds a combo chart, converts the second series to a line type, assigns a blue solid fill to the column series and a red stroke to the line series, then saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A – Categories
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Column B – Values for Column series
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(200);

            // Column C – Values for Line series
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(60);
            sheet.Cells["C5"].PutValue(80);

            // Add a Combo chart (initially a Column chart)
            int chartIdx = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIdx];

            // Add the column series (Sales)
            chart.NSeries.Add("B2:B5", true);
            // Add the line series (Profit)
            chart.NSeries.Add("C2:C5", true);

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A5";

            // Change the second series to a Line chart type
            chart.NSeries[1].Type = ChartType.Line;

            // Apply distinct fill colors
            // Column series – solid fill color
            chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(79, 129, 189); // blue shade

            // Line series – line color
            chart.NSeries[1].SeriesLines.Color = Color.FromArgb(192, 80, 77); // red shade

            // Save the workbook
            workbook.Save("ComboChartDistinctColors.xlsx");
        }
    }
}
