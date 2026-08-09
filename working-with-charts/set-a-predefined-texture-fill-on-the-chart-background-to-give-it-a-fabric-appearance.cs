// Title: Set a Fabric‑Like (Denim) Texture on a Chart Background with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add sample data, insert a column chart, and apply a predefined Denim texture to the chart area using Aspose.Cells' FillFormat (FillType.Texture, TextureType.Denim) before saving the file.
// Keywords: Aspose.Cells chart texture | C# chart background texture | TextureType.Denim | FillFormat FillType.Texture | Excel chart styling Aspose | predefined texture fill | .NET chart area fill | Aspose.Cells examples | chart background fabric appearance
// Common Searches: Aspose.Cells set denim texture on chart background | C# apply fabric texture to Excel chart area | How to use TextureType.Denim with Aspose.Cells | Chart background texture fill Aspose.Cells .NET | Predefined texture fill for charts in Aspose.Cells
// Developer Intent: Apply a predefined fabric texture to a chart's background area.
// Use Cases: Design a sales chart with a denim‑styled background to match brand guidelines. | Create presentation‑ready Excel charts that feature a fabric‑like visual effect. | Standardize the look of multiple charts in a workbook by applying the same texture fill.
// AI Prompts: Show how to change the chart background to another predefined texture such as Wool using Aspose.Cells. | Provide code for applying a custom image as a texture fill to the chart area instead of a built‑in texture. | Explain how to enumerate all available TextureType values programmatically in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add sample data, insert a column chart, and apply a predefined Denim texture to the chart area using Aspose.Cells' FillFormat (FillType.Texture, TextureType.Denim) before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
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

        // Apply a predefined texture (fabric-like) to the chart background
        chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;          // Use texture fill
        chart.ChartArea.Area.FillFormat.Texture = TextureType.Denim;          // Fabric appearance

        // Save the workbook
        workbook.Save("ChartWithFabricTexture.xlsx");
    }
}
