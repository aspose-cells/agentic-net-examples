// Title: C# – Aspose.Cells: Create a Chart Template with Fixed Legend Position and Data Labels
// Description: Demonstrates how to build a reusable Excel chart template in C# using Aspose.Cells: add sample data, insert a column chart, set the legend to the bottom, enable value data labels for the first series, customize label font size, and save the workbook as an .xlsx file.
// Keywords: Aspose.Cells | C# chart template | legend position bottom | data labels show value | column chart customization | SaveFormat.Xlsx | Excel automation | global chart template | US developers | Europe chart scripting
// Common Searches: Aspose.Cells set legend position bottom C# | How to enable data labels in Aspose.Cells chart | Create reusable chart template with Aspose.Cells | C# Aspose.Cells column chart with legend and labels | Save chart as template using Aspose.Cells
// Developer Intent: Generate a ready‑to‑use Excel chart template that enforces a specific legend placement and data‑label formatting for consistent reporting.
// Use Cases: Standardize sales or KPI charts across dozens of workbooks with a pre‑configured legend at the bottom. | Automate monthly reports where each column chart must display numeric values directly on the columns. | Distribute a template file to end‑users so they can replace data while preserving legend and label styles.
// AI Prompts: Write C# code with Aspose.Cells to create a line‑chart template that places the legend at the top and shows percentage data labels. | Provide a method that accepts a Workbook object and adds a bar chart with the legend on the right and a custom 14‑pt font for data labels. | Explain how to modify an existing Aspose.Cells chart template to move the legend to the left and hide all data labels.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTemplate
{
    // Demonstrates how to build a reusable Excel chart template in C# using Aspose.Cells: add sample data, insert a column chart, set the legend to the bottom, enable value data labels for the first series, customize label font size, and save the workbook as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet (topRow, leftColumn, bottomRow, rightColumn)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (including categories and values)
            chart.SetChartDataRange("A1:B4", true);

            // Predefine legend position (e.g., Bottom)
            chart.Legend.Position = LegendPositionType.Bottom;

            // Enable data labels to show values for the first series
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Optional: customize data label font size
            chart.NSeries[0].DataLabels.Font.Size = 12;

            // Save the workbook with the predefined chart template
            workbook.Save("ChartTemplateWithLegendAndDataLabels.xlsx", SaveFormat.Xlsx);
        }
    }
}
