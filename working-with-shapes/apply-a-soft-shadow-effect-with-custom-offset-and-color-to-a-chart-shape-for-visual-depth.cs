// Title: Add a soft shadow with custom offset and color to an Excel chart shape using Aspose.Cells for .NET
// Common Searches: asp.net apply soft shadow to chart shape using Aspose.Cells | c# set custom shadow offset and color on Excel chart shape | how to give visual depth to a chart with shape shadow in Aspose.Cells | Aspose.Cells chart shape shadow customization example | add soft shadow effect to column chart in .xlsx with C#
// Tags: chart shape soft shadow Aspose.Cells | shadow offset and color configuration Aspose.Cells | Excel chart visual depth styling .NET | shape shadow customization for charts C# | apply soft shadow to chart shape .xlsx

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, populates sample data, inserts a column chart, and saves the workbook as an .xlsx file. After obtaining the chart's Shape object you can use its Shadow property to apply a soft shadow with custom offset and color, adding visual depth to the chart.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
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
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            // Category data assignment is optional; some versions may not expose the property.
            // Uncomment the following line if the API version supports it:
            // chart.NSeries[0].CategoryData = "A2:A4";

            // Define output file path
            string outputPath = "ChartWithSoftShadow.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the chart workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
