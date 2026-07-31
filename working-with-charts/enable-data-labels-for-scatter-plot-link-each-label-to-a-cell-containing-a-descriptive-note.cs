// Title: Aspose.Cells for .NET – Add Linked Data Labels to a Scatter Chart (C#)
// Description: Creates a workbook, fills columns A‑C with X values, Y values, and notes, inserts a scatter chart, assigns X and Y ranges, enables data labels that pull text from the note cells (C2:C4), customizes label color and position, and saves the file as ScatterDataLabelsLinked.xlsx.
// Keywords: Aspose.Cells scatter chart | C# data labels linked to cells | Excel chart label source range | .NET chart customization | linked data labels Aspose | scatter plot label formatting
// Common Searches: Aspose.Cells link scatter chart labels to cells | C# set data label source range in Excel chart | How to customize scatter chart data labels Aspose.Cells | Enable ShowCellRange for chart series .NET | Add notes to scatter plot points using Aspose.Cells
// Developer Intent: Generate a scatter chart whose data labels display text from a worksheet range.
// Use Cases: Attach descriptive comments to each point in a scatter plot for analytical reports. | Create Excel workbooks where chart labels automatically reflect cell‑based notes that can be edited without changing code. | Apply consistent label styling (color, position) while linking to dynamic source data.
// AI Prompts: Show how to hide the Y value and display only the linked note in the scatter chart labels. | Provide code to add a second series with its own linked notes and distinct label formatting. | Explain how to update the LinkedSource range programmatically after appending new rows.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsScatterDataLabels
{
    // Creates a workbook, fills columns A‑C with X values, Y values, and notes, inserts a scatter chart, assigns X and Y ranges, enables data labels that pull text from the note cells (C2:C4), customizes label color and position, and saves the file as ScatterDataLabelsLinked.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate X values, Y values and descriptive notes
            // X values in column A, Y values in column B, notes in column C
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["C1"].PutValue("Note");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue("First point");

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue("Second point");

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["C4"].PutValue("Third point");

            // Add a scatter chart
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add series: Y values are the primary data source, X values are set separately
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].XValues = "A2:A4";

            // Enable data labels for the series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;               // Show the Y value
            series.DataLabels.ShowCellRange = true;           // Use cell range for label text
            series.DataLabels.LinkedSource = "C2:C4";         // Link each label to the note cell

            // Optional: adjust label appearance
            series.DataLabels.Font.Color = Color.DarkBlue;
            series.DataLabels.Position = LabelPositionType.Above;

            // Save the workbook
            workbook.Save("ScatterDataLabelsLinked.xlsx");
        }
    }
}
