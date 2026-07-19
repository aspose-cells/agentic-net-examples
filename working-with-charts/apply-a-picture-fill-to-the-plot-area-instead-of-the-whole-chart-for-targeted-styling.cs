// Title: C# – Apply a picture (texture) fill to a chart’s plot area using Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a workbook, adds sample data, inserts a column chart, and applies a picture (texture) fill only to the chart’s plot area. It sets the FillFormat to Texture, selects the Stretch picture mode, loads image data from a Base64 string, and saves the workbook as PlotAreaPictureFill.xlsx.
// Keywords: Aspose.Cells | C# chart plot area fill | picture texture fill Aspose.Cells | chart background image .NET | FillFormat Texture | plot area image fill | Aspose.Cells example | Excel chart styling
// Common Searches: Aspose.Cells set picture fill for chart plot area | C# apply texture background to Excel chart plot area | How to use FillFormat to add image to chart plot area in .NET | Aspose.Cells plot area background image example | Chart plot area picture fill Aspose.Cells
// Developer Intent: Add a picture or texture background to a chart’s plot area without changing the rest of the chart.
// Use Cases: Insert a company logo as a watermark inside the plot area while keeping the chart frame clean. | Use a custom texture to differentiate the data region from the surrounding chart elements in a presentation. | Replace the placeholder pixel with any PNG/JPEG to style the plot area for reports or dashboards.
// AI Prompts: Show how to load an external PNG file and apply it as a picture fill to a chart’s plot area with Aspose.Cells for .NET. | Provide code to change the picture fill mode from Stretch to Tile for a chart plot area. | Explain how to adjust the opacity of a plot area picture fill in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This Aspose.Cells for .NET example creates a workbook, adds sample data, inserts a column chart, and applies a picture (texture) fill only to the chart’s plot area. It sets the FillFormat to Texture, selects the Stretch picture mode, loads image data from a Base64 string, and saves the workbook as PlotAreaPictureFill.xlsx.
class PlotAreaPictureFillExample
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
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply picture fill to the plot area only
        // Set fill type to Texture (picture)
        chart.PlotArea.Area.FillFormat.FillType = FillType.Texture;
        // Choose picture fill mode (e.g., Stretch)
        chart.PlotArea.Area.FillFormat.PictureFormatType = FillPictureType.Stretch;
        // Provide image data – here we use a simple 1x1 white pixel as a placeholder
        string base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
        chart.PlotArea.Area.FillFormat.ImageData = Convert.FromBase64String(base64Image);

        // Save the workbook
        workbook.Save("PlotAreaPictureFill.xlsx");
    }
}
