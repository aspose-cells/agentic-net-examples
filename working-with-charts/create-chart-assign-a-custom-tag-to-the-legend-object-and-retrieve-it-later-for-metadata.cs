// Title: Store and Retrieve a Custom Tag in a Chart Legend with Aspose.Cells for .NET (C#)
// Description: This example shows how to create a workbook, add a column chart, embed a custom identifier in the legend using the Legend.Text property, save the file, reload it, and read the tag back—all without altering the chart's visual appearance.
// Keywords: Aspose.Cells | C# chart legend tag | custom metadata in Excel chart | Legend.Text property | store hidden data in chart | retrieve legend tag | column chart Aspose.Cells | Excel workbook metadata | Aspose.Cells API example | chart legend custom identifier
// Common Searches: How to add hidden metadata to a chart legend in Aspose.Cells C# | Retrieve custom tag from Excel chart legend after saving | Use Legend.Text to store data in Aspose.Cells chart | Aspose.Cells store and read custom information in chart objects | C# example for embedding identifiers in Excel chart legends
// Developer Intent: Embed a custom identifier in a chart legend and read it back after the workbook is saved.
// Use Cases: Link a chart to an external data source by storing a lookup key in the legend. | Record version or author information for audit trails without changing the chart appearance. | Enable dynamic formatting by keeping a hidden tag that drives runtime styling logic.
// AI Prompts: Generate C# code that saves a JSON payload as a custom tag in a chart legend using Aspose.Cells and retrieves it later. | Suggest an alternative way to hide metadata in a chart object with Aspose.Cells without using the visible legend text. | Create a sample that stores multiple custom tags across different chart elements and reads them after loading the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example shows how to create a workbook, add a column chart, embed a custom identifier in the legend using the Legend.Text property, save the file, reload it, and read the tag back—all without altering the chart's visual appearance.
class Program
{
    static void Main()
    {
        // ---------- Create a workbook and add sample data ----------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // ---------- Add a chart ----------
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // ---------- Assign a custom tag to the legend ----------
        // Aspose.Cells does not provide a dedicated Tag property for Legend.
        // As a workaround, we store the tag in the Legend's Text property.
        // This value can be read later without affecting the visual appearance
        // if the legend is hidden or its text is not displayed.
        string customLegendTag = "MyCustomLegendTag_2023";
        chart.Legend.Text = customLegendTag;

        // Optionally hide the legend text from being shown in the chart.
        // The legend itself can remain visible (ShowLegend = true) while its title text is not rendered.
        chart.ShowLegend = true;          // ensure legend is present
        chart.Legend.IsDeleted = false;   // keep legend object
        // The Text property is not displayed as a separate title, so the tag remains invisible.

        // ---------- Save the workbook ----------
        workbook.Save("ChartWithLegendTag.xlsx");

        // ---------- Later: Load the workbook and retrieve the tag ----------
        Workbook loadedWorkbook = new Workbook("ChartWithLegendTag.xlsx");
        Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
        Chart loadedChart = loadedSheet.Charts[0];

        // Retrieve the custom tag from the Legend's Text property
        string retrievedTag = loadedChart.Legend.Text;
        Console.WriteLine("Retrieved Legend Tag: " + retrievedTag);
    }
}
