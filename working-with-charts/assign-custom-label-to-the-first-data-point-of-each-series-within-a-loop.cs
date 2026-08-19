// Title: Aspose.Cells for .NET: Set a custom label on the first data point of each chart series (C#)
// Description: Demonstrates how to create a workbook, add a column chart with multiple series, loop through the series, make the first point's data label visible, and assign a custom text based on the series name using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# chart custom label | first data point label Aspose.Cells | modify chart series Aspose.Cells .NET | ChartPoint DataLabels ShowValue | loop through NSeries Aspose.Cells | Excel chart label example | Aspose.Cells chart programming
// Common Searches: Aspose.Cells set custom label for first point in chart series | C# loop through chart series and change data label Aspose.Cells | How to show value and custom text on first chart point using Aspose.Cells | Aspose.Cells .NET example modify chart point labels | Assign series name to first data point label in Excel chart
// Developer Intent: Add a unique text label to the first data point of every series in an Excel chart generated with Aspose.Cells.
// Use Cases: Highlight the opening sales figure of each product line in a column chart by displaying "First of <SeriesName>". | Mark the start date of multiple project phases in a timeline chart with a custom annotation on the first point of each series. | Create a financial report where baseline values for different accounts are emphasized with a custom label on the initial chart point.
// AI Prompts: Generate C# code with Aspose.Cells that adds a column chart and sets a custom data label on the first point of each series while hiding labels for the rest. | Explain how to retrieve the series name in Aspose.Cells and assign it to the first ChartPoint's DataLabels.Text property. | Show a loop over chart.NSeries that enables ShowValue only for the first point of each series and applies a custom label, leaving other points unchanged.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a column chart with multiple series, loop through the series, make the first point's data label visible, and assign a custom text based on the series name using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: categories in column A, two series in columns B and C
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add the two series (vertical orientation) and set category data
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Loop through each series in the chart
        foreach (Series series in chart.NSeries)
        {
            // Access the first data point of the current series
            ChartPoint firstPoint = series.Points[0];

            // Ensure the data label for this point is visible
            firstPoint.DataLabels.ShowValue = true;

            // Assign a custom label text to the first point
            // Example: "First of Series1" or "First of Series2"
            firstPoint.DataLabels.Text = $"First of {series.Name}";
        }

        // Optional: recalculate the chart to apply changes
        chart.Calculate();

        // Save the workbook to a file
        workbook.Save("CustomFirstPointLabels.xlsx");
    }
}
