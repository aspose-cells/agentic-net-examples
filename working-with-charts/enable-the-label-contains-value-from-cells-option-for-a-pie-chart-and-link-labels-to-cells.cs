// Title: Aspose.Cells for .NET – Set Pie Chart Data Labels from Worksheet Cells (C#)
// Description: Demonstrates how to create a workbook, populate category, value and label columns, add a pie chart, and configure the series so that data labels are taken from a cell range (ShowCellRange = true, LinkedSource = "C2:C4"). The example also hides the default numeric values, sets the label position to BestFit, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells pie chart data labels | C# ShowCellRange | LinkedSource chart labels | custom pie chart labels from cells | Aspose.Cells .NET chart example | Excel pie chart label from range
// Common Searches: Aspose.Cells enable label from cells pie chart | C# set pie chart data labels to cell range | ShowCellRange property Aspose.Cells | Link chart labels to worksheet cells .NET | Hide numeric values in Aspose.Cells pie chart
// Developer Intent: Configure a pie chart so its slice labels are sourced from worksheet cells instead of the default values.
// Use Cases: Display product colors or codes on each slice by linking to a column of descriptive text. | Create a sales‑by‑region chart where region names are stored separately and update automatically. | Build a reusable reporting template where changing label cells instantly refreshes chart labels without code changes.
// AI Prompts: Generate C# code with Aspose.Cells to create a doughnut chart whose data labels come from range D2:D5 and hide the numeric values. | Explain the role of ShowCellRange and LinkedSource when customizing chart data labels in Aspose.Cells. | Provide instructions to set the label position to InsideEnd for a bar chart while using cell‑based labels in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsPieChartLabelFromCells
{
    // Demonstrates how to create a workbook, populate category, value and label columns, add a pie chart, and configure the series so that data labels are taken from a cell range (ShowCellRange = true, LinkedSource = "C2:C4"). The example also hides the default numeric values, sets the label position to BestFit, and saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Cells that contain the label text to be shown on the pie slices
            sheet.Cells["C1"].PutValue("Label");
            sheet.Cells["C2"].PutValue("Red");
            sheet.Cells["C3"].PutValue("Yellow");
            sheet.Cells["C4"].PutValue("Red");

            // Add a pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values)
            chart.NSeries.Add("B2:B4", true);
            // Set the category (optional, not required for label-from-cells)
            chart.NSeries.CategoryData = "A2:A4";

            // Configure data labels to use values from cells
            Series series = chart.NSeries[0];
            series.DataLabels.ShowCellRange = true;          // Enable "Label Contains – Value From Cells"
            series.DataLabels.LinkedSource = "C2:C4";        // Link labels to the specified cell range
            series.DataLabels.ShowValue = false;            // Hide the default numeric value

            // Optional: adjust label position for better readability
            series.DataLabels.Position = LabelPositionType.BestFit;

            // Save the workbook
            workbook.Save("PieChartLabelFromCells.xlsx");
        }
    }
}
