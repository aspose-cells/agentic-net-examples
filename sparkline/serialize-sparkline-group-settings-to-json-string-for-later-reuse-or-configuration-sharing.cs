// Title: Serialize Aspose.Cells sparkline group configuration to a JSON string using C#
// AI Prompts: Generate C# code that creates a line sparkline group, customizes its visual properties, and saves the workbook to a JSON string with JsonSaveOptions. | Show how to export only the first worksheet containing a sparkline group to JSON using a MemoryStream in Aspose.Cells for .NET. | Provide a snippet that sets series, high‑point, and low‑point colors for a sparkline group and retrieves the resulting JSON representation.
// Common Searches: aspnet serialize sparkline group settings to JSON with Aspose.Cells | C# export worksheet containing sparklines to JSON string | how to use JsonSaveOptions to include sparkline colors in JSON output | save Aspose.Cells workbook as compact JSON without empty cells | retrieve JSON string of sparkline configuration from memory stream in .NET
// Tags: Aspose.Cells sparkline group JSON serialization | C# JsonSaveOptions export worksheet | sparkline group custom colors Aspose.Cells | export workbook to JSON string .NET | line sparkline configuration Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Json;

// The example creates a workbook, adds a line sparkline group for range A1:D1, configures high/low/first/last points, markers, line weight, and custom series, high‑point, and low‑point colors, then uses JsonSaveOptions to export only the first worksheet to a compact JSON string via a MemoryStream and prints the result.
class SparklineGroupJsonSerialization
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the location where the sparkline will be placed
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group (Line type) with the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Configure various sparkline group settings
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;
        group.ShowFirstPoint = true;
        group.ShowLastPoint = true;
        group.ShowMarkers = true;
        group.PlotRightToLeft = false;
        group.LineWeight = 1.5;

        // Set colors using CellsColor objects
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = System.Drawing.Color.Orange;
        group.SeriesColor = seriesColor;

        CellsColor highPointColor = workbook.CreateCellsColor();
        highPointColor.Color = System.Drawing.Color.Green;
        group.HighPointColor = highPointColor;

        CellsColor lowPointColor = workbook.CreateCellsColor();
        lowPointColor.Color = System.Drawing.Color.Red;
        group.LowPointColor = lowPointColor;

        // Serialize the workbook (including the sparkline group) to JSON string
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export only the first worksheet to keep the JSON concise
            SheetIndexes = new int[] { 0 },
            // Export as a JSON object even if there is only one sheet
            AlwaysExportAsJsonObject = true,
            // Do not export empty cells as null
            ExportEmptyCells = false,
            // No indentation for compact output (optional)
            Indent = ""
        };

        using (MemoryStream ms = new MemoryStream())
        {
            // Save the workbook to the memory stream using JSON format
            workbook.Save(ms, jsonOptions);

            // Convert the stream content to a UTF‑8 string
            string json = Encoding.UTF8.GetString(ms.ToArray());

            // Output the JSON string (could be stored, transmitted, etc.)
            Console.WriteLine(json);
        }
    }
}
