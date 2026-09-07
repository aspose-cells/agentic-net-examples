// Title: How to hide the legend of a column chart using Aspose.Cells for .NET to create a clean Excel presentation slide
// AI Prompts: Generate C# code that builds a column chart with Aspose.Cells and sets ShowLegend to false. | Write a C# snippet that removes the legend from an existing Aspose.Cells chart object. | Provide a complete Aspose.Cells .NET example that creates sample data, adds a column chart, disables its legend, and saves the workbook.
// Common Searches: aspnet hide legend column chart aspose.cells example | c# aspose.cells remove chart legend for presentation slides | how to disable legend in Excel chart using Aspose.Cells .NET | minimalist chart layout aspose.cells column chart without legend
// Tags: Aspose.Cells chart legend suppression | column chart without legend Aspose.Cells | Aspose.Cells ShowLegend property usage | minimalist Excel chart design Aspose.Cells | C# Aspose.Cells chart customization

using Aspose.Cells;
using Aspose.Cells.Charts;

// // Creates a workbook, adds sample data, inserts a column chart, disables its legend via ShowLegend = false, and saves the file as ChartWithoutLegend.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(15);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data series and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Hide the chart legend for a minimalist layout
        chart.ShowLegend = false;

        // Save the workbook with the chart
        workbook.Save("ChartWithoutLegend.xlsx");
    }
}
