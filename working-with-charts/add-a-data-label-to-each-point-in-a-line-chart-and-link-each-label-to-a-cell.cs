// Title: Aspose.Cells for .NET – Add Linked Data Labels to a Line Chart (C#)
// Description: This example creates an XLSX workbook, fills it with category, value, and label columns, inserts a line chart, enables data labels, positions them above each point, links each label to a cell in column C, synchronizes number formatting, customizes the font, and saves the file. It demonstrates how to bind chart labels to worksheet cells using Aspose.Cells APIs.
// Keywords: Aspose.Cells | C# line chart | linked data labels | ChartType.Line | Series.DataLabels | LinkedSource | LabelPositionType.Above | Excel chart sample | GitHub example | Aspose.Cells for .NET
// Common Searches: Aspose.Cells add data labels to line chart C# | link chart data labels to worksheet cells Aspose.Cells | set label position above points line chart Aspose | customize chart data label font Aspose.Cells | sample code Aspose.Cells line chart with linked labels
// Developer Intent: Generate a line chart with a data label for every point and bind each label to a specific worksheet cell.
// Use Cases: Display quarterly sales figures with custom period tags (Q1, Q2, Q3) directly on the chart. | Keep label formatting consistent by linking labels to source cells that may change. | Improve readability by positioning labels above points and applying a distinct font style.
// AI Prompts: Write C# code using Aspose.Cells that creates a line chart, shows data labels for each point, links the labels to a given cell range, and sets the label font color and size. | Explain how linked data labels update automatically when the underlying worksheet cells are edited in an Aspose.Cells workbook. | Provide step‑by‑step instructions to position data labels above line‑chart points and bind them to a different column using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsLineChartDataLabels
{
    // This example creates an XLSX workbook, fills it with category, value, and label columns, inserts a line chart, enables data labels, positions them above each point, links each label to a cell in column C, synchronizes number formatting, customizes the font, and saves the file. It demonstrates how to bind chart labels to worksheet cells using Aspose.Cells APIs.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A – Category (X‑axis)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // Column B – Values for the line series
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Column C – Text that will be linked to each data label
            sheet.Cells["C1"].PutValue("Label");
            sheet.Cells["C2"].PutValue("Q1");
            sheet.Cells["C3"].PutValue("Q2");
            sheet.Cells["C4"].PutValue("Q3");

            // Add a line chart
            int chartIndex = sheet.Charts.Add(ChartType.Line, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and the category (X) axis
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;               // Show the numeric value
            series.DataLabels.Position = LabelPositionType.Above; // Position appropriate for line charts

            // Link each data label to the corresponding cell in column C
            series.DataLabels.LinkedSource = "C2:C4";
            series.DataLabels.NumberFormatLinked = true;      // Keep number format in sync with source cells

            // Optional: customize appearance (font color, size, etc.)
            series.DataLabels.Font.Color = Color.DarkBlue;
            series.DataLabels.Font.Size = 10;

            // Save the workbook
            workbook.Save("LineChartWithLinkedDataLabels.xlsx");
        }
    }
}
