// Title: How to apply a triangle-shaped data label to a column chart series using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart and sets each point's data label marker to a triangle. | Show the Aspose.Cells API calls required to change a series' data label shape to a triangle in a .NET workbook. | Modify an existing Aspose.Cells chart example so that the data labels appear as triangles and save the workbook as an .xlsx file.
// Common Searches: aspnet aspose.cells change data label shape to triangle in column chart | c# set triangle marker for chart data labels using Aspose.Cells | how to customize data label marker type in Aspose.Cells chart series | Aspose.Cells .NET example for triangular data labels on column charts
// Tags: set triangular data label shape Aspose.Cells | column chart series data labels .NET | custom chart formatting Aspose.Cells | Aspose.Cells chart series marker type | apply triangle shape to chart labels

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, adds sample data, builds a column chart, enables value data labels, sets the data label shape to a triangle, and saves the file as CustomChart_TriangleDataLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet and give it a name
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Data";

            // Fill sample data for the chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];
            chart.Title.Text = "Custom Chart with Triangle Data Labels";

            // Add series and set category data
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Show data labels (value) on the chart
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Save the workbook
            string outputPath = "CustomChart_TriangleDataLabels.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
