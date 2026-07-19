// Title: Add Custom Text Labels from a Cell Range to an Aspose.Cells Column Chart (C#)
// Description: Demonstrates how to create a workbook, fill columns with categories, values and custom labels, insert a column chart, and configure the series so the data labels are taken from a separate cell range (C2:C4) instead of the default numeric values. The example also shows how to hide the numeric labels and position the custom text inside the column tops before saving the file.
// Keywords: Aspose.Cells C# chart custom labels | column chart data labels from cells | Series.DataLabels.LinkedSource | ShowCellRange Aspose.Cells | .NET Excel chart labeling | chart label cell range | Aspose.Cells chart example | custom text labels Excel chart | Aspose.Cells data label positioning | C# Excel chart API
// Common Searches: Aspose.Cells set data labels from a cell range | C# column chart custom labels Aspose.Cells | link chart data labels to cells .NET | hide numeric data labels Aspose.Cells chart | position data labels inside column Aspose.Cells
// Developer Intent: Replace the default numeric data labels on a column chart with custom text taken from a separate column of cells.
// Use Cases: Display descriptive names (e.g., Alpha, Beta, Gamma) on each column instead of numbers. | Allow non‑technical users to edit chart labels by simply changing cell values. | Create cleaner visualizations by hiding numeric values and showing only custom text inside the column ends.
// AI Prompts: Write C# code using Aspose.Cells to show data labels from a specified cell range on a column chart. | Explain how to hide numeric data labels and position custom text labels inside the top of each column in Aspose.Cells. | Provide an Aspose.Cells example that links a separate column of cells as data labels for a pie chart.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, fill columns with categories, values and custom labels, insert a column chart, and configure the series so the data labels are taken from a separate cell range (C2:C4) instead of the default numeric values. The example also shows how to hide the numeric labels and position the custom text inside the column tops before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate category names, values, and a separate column for custom labels
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Label");
        sheet.Cells["C2"].PutValue("Alpha");
        sheet.Cells["C3"].PutValue("Beta");
        sheet.Cells["C4"].PutValue("Gamma");

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set series values and category axis data
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure data labels to use the separate label column (C2:C4)
        Series series = chart.NSeries[0];
        series.DataLabels.ShowCellRange = true;          // Enable using a cell range for labels
        series.DataLabels.LinkedSource = "C2:C4";        // Link to the custom label column
        series.DataLabels.ShowValue = false;            // Hide the default numeric value
        series.DataLabels.Position = LabelPositionType.InsideEnd;

        // Save the workbook
        workbook.Save("ChartWithCustomLabels.xlsx", SaveFormat.Xlsx);
    }
}
