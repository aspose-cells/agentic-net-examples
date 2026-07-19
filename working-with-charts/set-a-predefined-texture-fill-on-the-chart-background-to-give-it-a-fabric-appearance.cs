// Title: Set a Denim Fabric Texture on a Chart Background with Aspose.Cells for .NET
// Description: Shows how to generate an Excel workbook, insert a column chart, and apply the predefined Denim texture to the chart area using Aspose.Cells for .NET. The example enables tiling so the pattern repeats and saves the workbook as an XLSX file.
// Keywords: Aspose.Cells chart texture | Denim texture chart background | C# FillType.Texture | TextureType.Denim example | Excel chart tiling | chart area fill Aspose.Cells | .NET Excel styling | GitHub Aspose.Cells chart texture sample
// Common Searches: Aspose.Cells apply denim texture to chart background C# | How to set chart area fill to texture in .NET | Enable tiling for chart background texture using Aspose.Cells | C# example of chart background fabric fill | Aspose.Cells texture fill types for charts
// Developer Intent: Apply a predefined fabric‑style texture to the background of an Excel chart.
// Use Cases: Brand‑specific sales reports that use a denim‑styled chart background for visual identity. | Automated dashboard generation where each chart shares a consistent fabric texture across worksheets. | Exporting charts to presentation‑ready files with a linen‑like background for a polished look.
// AI Prompts: Generate C# code with Aspose.Cells that sets a linen texture on a pie chart background, enables tiling, and saves the workbook as PDF. | List all values of the TextureType enum in Aspose.Cells and show how to let a user choose one for a chart area fill. | Create an example that applies a wood texture to a chart, disables tiling, and exports the result to XLSX.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Shows how to generate an Excel workbook, insert a column chart, and apply the predefined Denim texture to the chart area using Aspose.Cells for .NET. The example enables tiling so the pattern repeats and saves the workbook as an XLSX file.
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

        // Apply a predefined texture fill to the chart background (fabric appearance)
        chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;
        chart.ChartArea.Area.FillFormat.Texture = TextureType.Denim; // fabric-like texture
        // Enable tiling so the texture repeats across the background
        chart.ChartArea.Area.FillFormat.TextureFill.IsTiling = true;

        // Save the workbook with the textured chart
        workbook.Save("ChartWithFabricTexture.xlsx");
    }
}
