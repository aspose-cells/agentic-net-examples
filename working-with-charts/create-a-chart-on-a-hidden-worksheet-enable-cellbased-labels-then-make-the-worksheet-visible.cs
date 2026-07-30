// Title: Aspose.Cells for .NET – Create a Chart on a Hidden Worksheet, Apply Cell‑Based Data Labels, and Unhide the Sheet
// Description: Demonstrates how to add a column chart to a hidden worksheet, link its data labels to a cell range (column C) with custom font color, then change the worksheet visibility to visible and save the workbook as an XLSX file using Aspose.Cells in C#.
// Keywords: Aspose.Cells hidden worksheet chart | C# cell based data labels Aspose.Cells | Aspose.Cells set worksheet visibility | link chart data labels to cell range | Aspose.Cells chart label font color | create column chart Aspose.Cells .NET | unhide sheet after adding chart Aspose.Cells
// Common Searches: add chart to hidden sheet Aspose.Cells C# | enable cell based data labels Aspose.Cells | unhide worksheet after chart creation Aspose.Cells | set data label source range Aspose.Cells | customize chart label font color Aspose.Cells
// Developer Intent: Add a chart to a hidden worksheet, use cell values for the chart's data labels, customize label appearance, then make the worksheet visible before saving.
// Use Cases: Generate intermediate charts on hidden sheets during automated report building, then expose only the final sheets to users. | Create dashboards where each data point’s label is sourced from a separate column for richer descriptions. | Programmatically hide worksheets to keep workbook size low, add charts with cell‑linked labels, and reveal only the necessary sheets in the final output.
// AI Prompts: Show C# code that creates a pie chart on a hidden worksheet, links its data labels to a cell range, and then unhides the sheet using Aspose.Cells. | Provide an Aspose.Cells example that hides a sheet, adds a line chart with cell‑based labels, changes label font color, and saves the workbook as XLSX. | Explain how to set the visibility of a worksheet after adding a chart and how to link data labels to cells in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Demonstrates how to add a column chart to a hidden worksheet, link its data labels to a cell range (column C) with custom font color, then change the worksheet visibility to visible and save the workbook as an XLSX file using Aspose.Cells in C#.
class ChartOnHiddenSheetDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet that will hold the chart
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet ws = workbook.Worksheets[sheetIndex];
            ws.Name = "HiddenChartSheet";

            // Hide the worksheet
            ws.VisibilityType = VisibilityType.Hidden;

            // Populate sample data
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["C1"].PutValue("Label");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["C2"].PutValue("Ten");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["C3"].PutValue("Twenty");
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B4"].PutValue(30);
            ws.Cells["C4"].PutValue("Thirty");

            // Add a column chart to the hidden worksheet
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = ws.Charts[chartIdx];

            // Define the data series and category axis
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable cell‑based data labels (use values from column C)
            var series = chart.NSeries[0];
            series.DataLabels.ShowCellRange = true;
            series.DataLabels.LinkedSource = "C2:C4";
            series.DataLabels.Font.Color = Color.Blue;

            // Make the worksheet visible again
            ws.VisibilityType = VisibilityType.Visible;

            // Save the workbook
            workbook.Save("ChartOnHiddenSheet.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
