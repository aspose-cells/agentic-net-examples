// Title: Create a 3‑D Column chart and customize its perspective depth and rotation angles using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set perspective depth for a 3D column chart | How to change rotation X and Y angles of a Column3D chart in .NET | Adjust 3D chart view angle programmatically with Aspose.Cells | Customize depth and elevation of Excel 3D column chart using C# | Set chart perspective and rotation properties for Column3D in Aspose.Cells
// Tags: 3d column chart perspective Aspose.Cells | chart rotation angle .NET | adjust chart depth C# | column3d view customization Aspose.Cells | excel 3d chart formatting C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new Workbook, writes category and value data to cells A1:B4, adds a 3‑D Column chart (ChartType.Column3D) to the worksheet, links the series to the range A2:B4, and saves the file as 3DColumnChart.xlsx. After the chart is created you can modify its 3‑D appearance—such as perspective depth, RotationX, and RotationY—through the Chart.Perspective, Chart.RotationX and Chart.RotationY properties.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a 3‑D column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and categories together
            // "A2:A4" are categories, "B2:B4" are values
            chart.NSeries.Add("A2:B4", true);

            // Save the workbook
            string outputPath = "3DColumnChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
