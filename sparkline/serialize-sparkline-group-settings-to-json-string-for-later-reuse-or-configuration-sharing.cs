// Title: C# – Serialize Aspose.Cells SparklineGroup Settings to JSON
// Description: Demonstrates how to create a workbook, add a sparkline group, customize its properties (high/low points, line weight, preset style, colors, etc.), and export the entire worksheet—including the sparkline configuration—to a formatted JSON string using JsonSaveOptions and a memory stream.
// Keywords: Aspose.Cells | SparklineGroup | JSON serialization | C# | JsonSaveOptions | export sparkline settings | save sparkline configuration | Excel sparkline JSON | Aspose.Cells API | sparkline group properties
// Common Searches: Aspose.Cells serialize sparkline group to JSON | C# export sparkline settings as JSON | JsonSaveOptions example for sparkline | how to save Aspose.Cells sparkline configuration | convert Excel sparkline group to JSON string
// Developer Intent: Generate a JSON string that captures a SparklineGroup's configuration for later reuse or sharing.
// Use Cases: Persist custom sparkline formatting across workbook versions by storing the group settings in JSON. | Distribute a standard sparkline style to multiple reports by applying a saved JSON configuration. | Save sparkline parameters in a database or config file to dynamically render charts in generated Excel files.
// AI Prompts: Write C# code that reads the JSON produced by JsonSaveOptions and rebuilds the SparklineGroup in a new workbook using Aspose.Cells. | Show how to modify the line weight and colors directly in the exported JSON before re‑importing it. | Explain how to extract only the SparklineGroup section from the worksheet JSON without loading the full workbook.

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Json;
using System;
using System.Drawing;
using System.IO;
using System.Text;

// Demonstrates how to create a workbook, add a sparkline group, customize its properties (high/low points, line weight, preset style, colors, etc.), and export the entire worksheet—including the sparkline configuration—to a formatted JSON string using JsonSaveOptions and a memory stream.
class SparklineGroupJsonSerialization
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
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

        // Add a sparkline group and a sparkline inside it
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Configure various group settings that we want to serialize
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;
        group.LineWeight = 1.5;
        group.PresetStyle = SparklinePresetStyleType.Style3;
        group.Type = SparklineType.Column;
        group.PlotRightToLeft = false;
        group.DisplayHidden = true;

        // Set colors using CellsColor objects
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.Orange;
        group.SeriesColor = seriesColor;

        CellsColor highPointColor = workbook.CreateCellsColor();
        highPointColor.Color = Color.Green;
        group.HighPointColor = highPointColor;

        // Prepare JSON save options – export the whole sheet with indentation
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportArea = new CellArea { StartRow = 0, EndRow = 0, StartColumn = 0, EndColumn = 5 },
            ExportAsString = true,
            Indent = "  "
        };

        // Save the workbook to a memory stream as JSON and retrieve the string
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, jsonOptions);
            string json = Encoding.UTF8.GetString(ms.ToArray());

            // The JSON string now contains the sparkline group configuration
            Console.WriteLine(json);
        }
    }
}
