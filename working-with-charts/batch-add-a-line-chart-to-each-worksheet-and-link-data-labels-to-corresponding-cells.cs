// Title: Add line charts to multiple worksheets with cell‑linked data labels using Aspose.Cells for .NET (C#)
// Description: C# example that creates a new workbook, removes the default sheet, adds three worksheets with sample data, inserts a Line chart on each sheet, sets the series values from column B, uses column A for categories, links data labels to the formatted text in column C, positions the labels outside the points, and saves the file as XLSX.
// Keywords: Aspose.Cells line chart C# | batch chart creation Aspose.Cells | link chart data labels to cells | multiple worksheets chart Aspose | set data label position Aspose.Cells | ChartType.Line example | Aspose.Cells .NET tutorial
// Common Searches: how to add a line chart to every worksheet with Aspose.Cells | Aspose.Cells link data labels to a cell range | batch create charts in C# using Aspose.Cells | set data label position outside end Aspose.Cells chart | generate workbook with charts on each sheet Aspose
// Developer Intent: Create a line chart on each worksheet and bind its data labels to a corresponding cell range.
// Use Cases: Produce a sales performance workbook where each region sheet contains a line chart with labels showing formatted revenue figures from a separate column. | Automate a departmental KPI dashboard that adds a sheet per department, inserts a line chart, and links labels to descriptive notes for clearer reporting. | Generate a product testing report with multiple sheets, each featuring a line chart whose labels are linked to custom unit strings for precise visualization.
// AI Prompts: Write C# code with Aspose.Cells to add a bar chart to every worksheet and link its data labels to column D. | Modify the example to set each chart title to the worksheet name and export the workbook to PDF while keeping the linked labels. | Explain how to change the data label position to InsideEnd for all charts created in a batch using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that creates a new workbook, removes the default sheet, adds three worksheets with sample data, inserts a Line chart on each sheet, sets the series values from column B, uses column A for categories, links data labels to the formatted text in column C, positions the labels outside the points, and saves the file as XLSX.
class BatchLineChartWithLinkedLabels
{
    static void Main()
    {
        try
        {
            // Create a new workbook and clear the default sheet
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();

            // Add three worksheets with sample data and a line chart each
            for (int i = 0; i < 3; i++)
            {
                Worksheet ws = workbook.Worksheets.Add($"Sheet{i + 1}");

                // Header row
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["C1"].PutValue("Label");

                // Populate rows 2‑11
                for (int row = 2; row <= 11; row++)
                {
                    ws.Cells[$"A{row}"].PutValue($"Item {row - 1}");
                    int val = row * 10;
                    ws.Cells[$"B{row}"].PutValue(val);
                    ws.Cells[$"C{row}"].PutValue($"{val} units");
                }

                // Add a line chart
                int chartIdx = ws.Charts.Add(ChartType.Line, 5, 0, 20, 8);
                Chart chart = ws.Charts[chartIdx];

                // Set the values series (B2:B11) – the second argument indicates that the data is vertical
                chart.NSeries.Add($"={ws.Name}!$B$2:$B$11", true);

                // Set category (X‑axis) data from column A
                Series series = chart.NSeries[0];
                series.XValues = $"={ws.Name}!$A$2:$A$11";

                // Link data labels to the formatted cells in column C
                series.DataLabels.ShowCellRange = true;
                series.DataLabels.LinkedSource = $"={ws.Name}!$C$2:$C$11";
                series.DataLabels.Position = LabelPositionType.OutsideEnd;
            }

            // Save the workbook
            workbook.Save("BatchLineCharts.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
