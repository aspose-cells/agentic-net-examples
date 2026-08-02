// Title: Aspose.Cells C# Example: Hide Legend on a 3‑D Pie Chart
// Description: Shows how to build a workbook, insert fruit data, create a 3‑D pie chart, turn off its legend with chart.ShowLegend = false, add a descriptive title, and save the file. Highlights the impact on chart readability when the legend is omitted.
// Keywords: Aspose.Cells | C# | 3D pie chart | hide legend | chart.ShowLegend | Excel chart example | remove chart legend | chart readability | Aspose.Cells for .NET | chart title without legend
// Common Searches: hide legend Aspose.Cells 3D pie chart C# | remove chart legend Aspose.Cells example | Aspose.Cells hide legend and keep title | C# code to disable legend on pie chart | impact of hiding legend on Excel chart readability
// Developer Intent: The developer wants to suppress the legend of a 3‑D pie chart while keeping the visual information clear.
// Use Cases: Design compact dashboard charts where the title replaces the legend to save space. | Generate printable Excel reports that avoid legend overlap with data labels. | Create automated reporting scripts that produce clean pie charts for email attachments.
// AI Prompts: Provide C# code to hide the legend of a 3‑D pie chart using Aspose.Cells and add a title. | Explain how disabling chart.ShowLegend affects the layout of a 3‑D pie chart and suggest alternatives for data identification. | Generate a function that toggles the legend visibility of any Aspose.Cells chart based on a boolean parameter.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// Shows how to build a workbook, insert fruit data, create a 3‑D pie chart, turn off its legend with chart.ShowLegend = false, add a descriptive title, and save the file. Highlights the impact on chart readability when the legend is omitted.
class HideLegend3DPieChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pie chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["B3"].PutValue(45);
        sheet.Cells["B4"].PutValue(25);

        // Add a 3‑D pie chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Pie3D, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data series and categories for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Hide the legend to see how it affects readability
        chart.ShowLegend = false;

        // Add a title so the chart remains self‑explanatory without a legend
        chart.Title.Text = "Fruit Distribution";

        // Write the current legend visibility to the console (should be false)
        Console.WriteLine("Legend visible: " + chart.ShowLegend);

        // Save the workbook containing the chart
        workbook.Save("3DPieChart_NoLegend.xlsx");
    }
}
