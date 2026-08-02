// Title: Enable Data Labels (ShowValue) for First Series of a Column Chart – Aspose.Cells C# Example
// Description: C# snippet that creates a workbook, adds category and value data, inserts a column chart, defines the series range, activates data labels for the first series with Series.DataLabels.ShowValue = true, recalculates the chart, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | column chart | data labels | ShowValue | chart series values | Excel chart API | sample code | GitHub
// Common Searches: Aspose.Cells show data labels first series | C# enable point values on column chart Aspose.Cells | Series.DataLabels.ShowValue example | display values on Excel chart using Aspose.Cells | how to turn on data labels in Aspose.Cells chart
// Developer Intent: Turn on data labels for the first series of a chart so each point’s numeric value appears directly on the chart.
// Use Cases: Show exact sales numbers on each column of a quarterly sales chart. | Display measured temperatures on a column chart for a scientific report. | Add numeric labels to a budget allocation chart for a financial presentation.
// AI Prompts: Write C# code with Aspose.Cells that creates a bar chart and enables data labels for every series. | Provide an Aspose.Cells example that toggles data label visibility based on a runtime condition. | Explain how to customize font style, color, and position of data labels for a chart series using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// C# snippet that creates a workbook, adds category and value data, inserts a column chart, defines the series range, activates data labels for the first series with Series.DataLabels.ShowValue = true, recalculates the chart, and saves the file.
class EnableDataLabelsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Insert a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the series and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series to display point values
        Series firstSeries = chart.NSeries[0];
        firstSeries.DataLabels.ShowValue = true; // rule: enable data labels

        // Recalculate the chart (optional but ensures labels are updated)
        chart.Calculate();

        // Save the workbook
        workbook.Save("EnableDataLabelsDemo.xlsx");
    }
}
