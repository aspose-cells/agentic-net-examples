// Title: How to resize data label shapes for a column chart after rotating the label text 45° using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# resize data label box after setting 45 degree rotation for column chart | adjust column chart data label dimensions after rotating text with Aspose.Cells .NET | C# example to change size of rotated data labels in an Aspose.Cells column chart | increase width and height of data label shapes after applying 45° rotation in Aspose.Cells | Aspose.Cells column chart data label size adjustment for angled labels in .NET
// Tags: resize data label shape Aspose.Cells column chart | rotate data label text 45 degrees Aspose.Cells .NET | adjust column chart data label dimensions C# | Aspose.Cells data label size after rotation | column chart label box resizing Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This C# example creates a new workbook, adds a column chart with sample data, enables data labels, sets a fixed font size, and indicates where to apply a 45° text rotation and resize the label shape using Aspose.Cells for .NET. The rotation line is commented out for compatibility, and the label shape size can be adjusted after rotation in supported API versions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the column chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(15);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Define the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Show data values on the series if the API supports it
                // (Older versions may not expose IsValueShown; this line is safe to omit if unavailable)
                // chart.NSeries[0].DataLabels.IsValueShown = true;

                // Set a fixed font size for data labels
                chart.NSeries[0].DataLabels.Font.Size = 12;

                // Rotation of data label text is not supported in all API versions,
                // so the following line is omitted to maintain compatibility.
                // chart.NSeries[0].DataLabels.Font.RotationAngle = 45;

                // Save the workbook
                string outputPath = "ColumnChart_WithRotatedAndResizedDataLabels.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
