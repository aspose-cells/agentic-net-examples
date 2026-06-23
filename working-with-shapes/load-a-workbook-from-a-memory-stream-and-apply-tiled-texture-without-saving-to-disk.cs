using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a source workbook, add data and a chart.
        // ------------------------------------------------------------
        Workbook sourceWb = new Workbook();
        Worksheet ws = sourceWb.Worksheets[0];

        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["A2"].PutValue("A");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["A3"].PutValue("B");
        ws.Cells["B3"].PutValue(20);
        ws.Cells["A4"].PutValue("C");
        ws.Cells["B4"].PutValue(30);

        // Add a column chart that uses the data.
        int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 10);
        Chart chart = ws.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // ------------------------------------------------------------
        // 2. Save the workbook to a memory stream (no file I/O).
        // ------------------------------------------------------------
        using (MemoryStream ms = new MemoryStream())
        {
            // Rule: Workbook.Save(Stream, SaveFormat)
            sourceWb.Save(ms, SaveFormat.Xlsx);
            ms.Position = 0; // Reset stream for reading.

            // ------------------------------------------------------------
            // 3. Load the workbook from the memory stream.
            // ------------------------------------------------------------
            // Rule: Workbook(Stream) constructor
            Workbook wb = new Workbook(ms);

            // ------------------------------------------------------------
            // 4. Apply a tiled texture to the chart's plot area.
            // ------------------------------------------------------------
            Worksheet loadedWs = wb.Worksheets[0];
            Chart loadedChart = loadedWs.Charts[0];

            // Set the fill type to texture.
            loadedChart.PlotArea.Area.FillFormat.FillType = FillType.Texture;

            // Choose a built‑in texture (optional, but demonstrates usage).
            loadedChart.PlotArea.Area.FillFormat.Texture = TextureType.BlueTissuePaper;

            // Rule: TextureFill.IsTiling property
            loadedChart.PlotArea.Area.FillFormat.TextureFill.IsTiling = true;

            // Verify that the property was applied.
            Console.WriteLine("IsTiling set to: " + loadedChart.PlotArea.Area.FillFormat.TextureFill.IsTiling);
        }
    }
}