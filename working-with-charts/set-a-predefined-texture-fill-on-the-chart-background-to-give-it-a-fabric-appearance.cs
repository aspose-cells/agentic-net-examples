using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartTextureDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply a texture fill to the chart background (fabric appearance)
            // First set the fill type to Texture
            chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;
            // Then assign a predefined texture, e.g., Denim, which resembles fabric
            chart.ChartArea.Area.FillFormat.Texture = TextureType.Denim;
            // Alternatively, you can also set the TextureFill.Type property
            // chart.ChartArea.Area.FillFormat.TextureFill.Type = TextureType.Denim;

            // Save the workbook
            workbook.Save("ChartWithFabricTexture.xlsx");
        }
    }
}