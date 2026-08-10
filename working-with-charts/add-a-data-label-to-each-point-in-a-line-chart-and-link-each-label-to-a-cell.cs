// Title: Add Linked Data Labels to Each Point of a Line Chart with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills columns with categories, values and custom text, adds a line chart, enables data labels for the series, then links each point's label to the corresponding cell in column C, applies the source cell's number format, positions the label above the point, and saves the file.
// Keywords: Aspose.Cells line chart data labels | C# link chart label to cell | Aspose.Cells LinkedSource property | chart point label position above | set number format linked Aspose.Cells | add data labels to chart series | Aspose.Cells chart example C#
// Common Searches: How to link chart data labels to worksheet cells using Aspose.Cells | Aspose.Cells C# add data labels to each point of a line chart | Set label position above points in Aspose.Cells line chart | Link data label to cell and keep number format in Aspose.Cells | Create line chart with custom labels from a column in C#
// Developer Intent: Generate a line chart, enable data labels for every point, and bind each label to a specific worksheet cell.
// Use Cases: Sales trend line chart where each point shows a performance tag stored in a separate column. | Project timeline chart with milestone descriptions linked to data labels for automatic updates. | Financial line chart that mirrors the source cell's number formatting in its data labels.
// AI Prompts: Write C# code with Aspose.Cells to add a line chart, turn on data labels for each point, and link each label to a cell in column C, preserving the cell's number format and placing the label above the point. | Explain how to use the LinkedSource property to bind chart point data labels to worksheet cells and adjust label positioning in Aspose.Cells. | Provide a step‑by‑step tutorial for creating a line chart with custom text labels sourced from a range, including saving the workbook as an .xlsx file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills columns with categories, values and custom text, adds a line chart, enables data labels for the series, then links each point's label to the corresponding cell in column C, applies the source cell's number format, positions the label above the point, and saves the file.
class AddDataLabelsToLineChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        // Column A: Categories
        // Column B: Values for the line chart
        // Column C: Text to be linked to each data label
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["C1"].PutValue("Label");
        string[] categories = { "Jan", "Feb", "Mar", "Apr" };
        double[] values = { 10, 20, 15, 25 };
        string[] labels = { "Low", "Medium", "Medium-High", "High" };

        for (int i = 0; i < categories.Length; i++)
        {
            int row = i + 2; // Data starts from row 2
            sheet.Cells[$"A{row}"].PutValue(categories[i]);
            sheet.Cells[$"B{row}"].PutValue(values[i]);
            sheet.Cells[$"C{row}"].PutValue(labels[i]);
        }

        // Add a line chart
        int chartIndex = sheet.Charts.Add(ChartType.Line, 6, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the series
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Enable data labels for the series (required to access point labels)
        chart.NSeries[0].DataLabels.ShowValue = true;

        // Iterate through each point and link its data label to the corresponding cell in column C
        Series series = chart.NSeries[0];
        for (int i = 0; i < series.Points.Count; i++)
        {
            ChartPoint point = series.Points[i];
            // Show the value (optional, can be turned off if only linked text is needed)
            point.DataLabels.ShowValue = true;

            // Link the data label to the cell in column C of the same row
            string linkedCell = $"C{i + 2}";
            point.DataLabels.LinkedSource = linkedCell;

            // Ensure the label reflects the cell's number format (if any)
            point.DataLabels.NumberFormatLinked = true;

            // Position the label above the point (typical for line charts)
            point.DataLabels.Position = LabelPositionType.Above;
        }

        // Save the workbook
        workbook.Save("LineChartWithLinkedDataLabels.xlsx");
    }
}
