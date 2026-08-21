// Title: Enable Text Wrapping for Pie Chart Data Labels in C# with Aspose.Cells
// Description: This C# example creates a workbook, adds sales data, inserts a pie chart, shows values and category names on the labels, activates text wrapping via DataLabels.IsTextWrapped, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | pie chart | data labels | IsTextWrapped | text wrap | chart label formatting | Excel automation | sample code | GitHub example
// Common Searches: Aspose.Cells wrap text in pie chart labels | DataLabels.IsTextWrapped C# example | how to enable label wrapping in Excel chart using Aspose | pie chart data label multiline Aspose.Cells .NET | C# code for wrapping chart labels in Excel
// Developer Intent: Apply multiline wrapping to pie‑chart data labels so long category names and values are fully visible.
// Use Cases: Generate sales dashboards where category titles exceed label width. | Automate Excel reports that include pie charts with readable, wrapped labels. | Create templates for recurring exports that need consistent label formatting across multiple charts.
// AI Prompts: Provide C# code that sets DataLabels.IsTextWrapped = true for a pie chart using Aspose.Cells. | Show how to configure data label visibility and text wrapping for Excel charts in .NET. | Explain steps to apply multiline label wrapping to different chart types with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example creates a workbook, adds sales data, inserts a pie chart, shows values and category names on the labels, activates text wrapping via DataLabels.IsTextWrapped, and saves the file as an Excel workbook.
class EnableDataLabelWrapping
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the pie chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(85);
        sheet.Cells["B4"].PutValue(65);

        // Insert a pie chart into the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Access the data labels of the first series
        DataLabels labels = chart.NSeries[0].DataLabels;

        // Show values and category names on the data labels
        labels.ShowValue = true;
        labels.ShowCategoryName = true;

        // Enable text wrapping for the data labels
        labels.IsTextWrapped = true;

        // Save the workbook to a file
        workbook.Save("PieChartDataLabelsWrapped.xlsx");
    }
}
