// Title: Add a soft shadow with custom offset and color to an Excel chart shape using Aspose.Cells for .NET (unsupported)
// AI Prompts: Generate C# code that attempts to set a soft shadow with a specific offset and color on an Aspose.Cells chart shape, including error handling for unsupported features. | Suggest a C# technique to emulate a soft shadow effect on an Excel chart created with Aspose.Cells by using additional shapes or formatting tricks.
// Common Searches: asp.net aspose.cells add soft shadow to chart shape | c# aspose.cells chart shadow offset color not supported | how to mimic shadow on Excel chart using Aspose.Cells .NET | apply custom shadow to chart object programmatically with Aspose.Cells | visual depth for Excel chart Aspose.Cells workaround
// Tags: apply soft shadow to chart shape Aspose.Cells | custom shadow offset and color C# Aspose.Cells | chart shape styling Aspose.Cells .NET | simulate chart shadow Aspose.Cells workaround | Excel chart visual depth Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data and a column chart, and documents that Aspose.Cells for .NET does not currently support applying a soft shadow directly to chart shapes, prompting developers to consider alternative approaches.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            var chart = sheet.Charts[chartIndex];

            // Set the data source for the chart (values) and categories
            chart.NSeries.Add("B2:B4", true);          // values range, vertical series
            chart.NSeries.CategoryData = "A2:A4";      // categories range

            // NOTE: Shadow effect on charts is not available in the current Aspose.Cells version.
            // If needed, you can apply shadow to shapes or other objects that support it.

            // Save the workbook with the chart
            string outputPath = "ChartWithSoftShadow.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
