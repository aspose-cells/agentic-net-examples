// Title: Load a Workbook from MemoryStream and Enumerate Its ChartCollection with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a column chart, save it to a MemoryStream, reload the workbook from that stream, access the first worksheet's ChartCollection, and output each chart's type using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | .NET | MemoryStream | load workbook from stream | ChartCollection | enumerate charts | read Excel chart types | process Excel charts | Aspose.Cells example
// Common Searches: Aspose.Cells load workbook from MemoryStream | How to read chart collection after loading Excel file with Aspose.Cells | C# enumerate charts in a workbook using Aspose.Cells | Get chart type from Excel file with Aspose.Cells .NET | Access ChartCollection from stream in Aspose.Cells
// Developer Intent: Load an Excel workbook from a MemoryStream and iterate through its charts for inspection or further processing.
// Use Cases: List all chart types in a workbook loaded from a stream. | Modify chart properties (title, colors, data range) after loading from a MemoryStream. | Export each chart to an image format (PNG, JPEG) after stream‑based loading. | Validate that required charts exist before saving the workbook.
// AI Prompts: Write C# code using Aspose.Cells to load a workbook from a MemoryStream and print each chart's type. | Show how to change the title of every chart after loading an Excel file from a stream with Aspose.Cells for .NET. | Provide an example that loads a workbook from a MemoryStream and saves each chart as a PNG image using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a column chart, save it to a MemoryStream, reload the workbook from that stream, access the first worksheet's ChartCollection, and output each chart's type using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a workbook, add data and a chart (sample source file)
        // ------------------------------------------------------------
        Workbook sourceWb = new Workbook();
        Worksheet ws = sourceWb.Worksheets[0];

        ws.Cells["A1"].PutValue("Category");
        ws.Cells["A2"].PutValue("A");
        ws.Cells["A3"].PutValue("B");
        ws.Cells["A4"].PutValue("C");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["B3"].PutValue(20);
        ws.Cells["B4"].PutValue(30);

        // Add a column chart and bind it to the data range
        int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = ws.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // ------------------------------------------------------------
        // 2. Save the workbook to a memory stream (uses SaveToStream rule)
        // ------------------------------------------------------------
        MemoryStream ms = sourceWb.SaveToStream();

        // Reset the stream position before reading
        ms.Position = 0;

        // ------------------------------------------------------------
        // 3. Load a workbook from the memory stream (uses Workbook(Stream) rule)
        // ------------------------------------------------------------
        Workbook loadedWb = new Workbook(ms);

        // ------------------------------------------------------------
        // 4. Access the chart collection of the first worksheet
        // ------------------------------------------------------------
        Worksheet loadedWs = loadedWb.Worksheets[0];
        ChartCollection charts = loadedWs.Charts;

        // Example processing: list each chart's type
        for (int i = 0; i < charts.Count; i++)
        {
            Chart c = charts[i];
            Console.WriteLine($"Chart {i} type: {c.Type}");
        }

        // Clean up the memory stream
        ms.Dispose();
    }
}
