// Title: Add linked data labels to a scatter chart using Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, fills columns A‑C with X values, Y values and descriptive notes, builds a scatter chart, assigns XValues and YValues, enables data labels to display the Y value and to pull label text from the note range (C2:C5), sets the label position to the right, and saves the file as ScatterDataLabelsLinked.xlsx.
// Keywords: Aspose.Cells | C# | scatter chart | data labels | linked source | ShowCellRange | LabelPositionType | Excel automation | chart series | custom label text
// Common Searches: Aspose.Cells link data labels to cells | scatter chart data labels C# Aspose.Cells | set XValues for scatter chart Aspose.Cells | custom label text for each point Aspose.Cells | change data label position scatter chart .NET
// Developer Intent: I need to show a note from a worksheet cell as the data label for each point in a scatter chart.
// Use Cases: Generate an Excel report where each scatter point displays a comment stored in a separate column. | Create a scientific chart with inline annotations linked to source cells. | Automate workbook creation with custom positioned data labels for better readability. | Export data visualizations with linked labels for downstream analysis.
// AI Prompts: Write C# code with Aspose.Cells to create a scatter chart, set XValues from A2:A5, YValues from B2:B5, and link data labels to notes in C2:C5. | Show how to enable ShowCellRange and assign LinkedSource for series data labels in a scatter plot using Aspose.Cells for .NET. | Provide an example that changes the data label position to the right side of each point in a scatter chart with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsScatterDataLabels
{
    // This C# example creates a workbook, fills columns A‑C with X values, Y values and descriptive notes, builds a scatter chart, assigns XValues and YValues, enables data labels to display the Y value and to pull label text from the note range (C2:C5), sets the label position to the right, and saves the file as ScatterDataLabelsLinked.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate X values, Y values and descriptive notes
            // X values in column A, Y values in column B, notes in column C
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["C1"].PutValue("Note");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["A5"].PutValue(4);

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["B5"].PutValue(25);

            sheet.Cells["C2"].PutValue("First point");
            sheet.Cells["C3"].PutValue("Second point");
            sheet.Cells["C4"].PutValue("Third point");
            sheet.Cells["C5"].PutValue("Fourth point");

            // Add a scatter chart
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set Y values range; X values are set via XValues property
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].XValues = "A2:A5";

            // Enable data labels and link each label to the corresponding note cell
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;          // Show the Y value
            series.DataLabels.ShowCellRange = true;      // Use cell range for label text
            series.DataLabels.LinkedSource = "C2:C5";    // Range containing notes

            // Optional: adjust label position if desired
            series.DataLabels.Position = LabelPositionType.Right;

            // Save the workbook
            workbook.Save("ScatterDataLabelsLinked.xlsx");
        }
    }
}
