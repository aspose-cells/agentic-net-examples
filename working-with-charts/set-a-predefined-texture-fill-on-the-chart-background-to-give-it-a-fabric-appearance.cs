// Title: Apply a fabric‑style texture fill to a chart’s background using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets the ChartArea’s FillFormat.Texture to TextureType.WovenMat and enables tiling with Aspose.Cells. | Demonstrate how to add a column chart, apply a predefined fabric texture to its background, and save the workbook using Aspose.Cells.
// Common Searches: c# aspose.cells set chartarea texture fill to woven mat | how to enable tiling for chart background texture in Aspose.Cells | apply predefined fabric texture to Excel chart using Aspose.Cells .NET | example of using TextureType.WovenMat for chart background in C# | chart area fill format texture options aspose.cells tutorial
// Tags: chartarea texture fill aspose.cells | wovenmat texture aspose.cells | chart background tiling aspose.cells | excel chart fabric texture c# | set chartarea fillformat texture aspose.cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data, inserts a column chart, applies the predefined WovenMat texture to the ChartArea with tiling enabled, and saves the file as ChartWithFabricTexture.xlsx.
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

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply a fabric‑like texture to the chart background (ChartArea)
        chart.ChartArea.Area.FillFormat.Texture = TextureType.WovenMat; // predefined fabric texture
        // Enable tiling so the texture repeats across the whole area
        chart.ChartArea.Area.FillFormat.TextureFill.IsTiling = true;

        // Save the workbook
        workbook.Save("ChartWithFabricTexture.xlsx");
    }
}
