using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Json;

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

        // Add a sparkline group with a line type
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Configure various group settings
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;
        group.ShowFirstPoint = true;
        group.ShowLastPoint = true;
        group.ShowMarkers = true;
        group.PlotRightToLeft = false;
        group.LineWeight = 1.5;
        group.PresetStyle = SparklinePresetStyleType.Style3;
        group.Type = SparklineType.Line;

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

        // Add a sparkline to the group (required for the group to be valid)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Prepare JSON save options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Export the whole workbook as JSON; the sparkline group settings will be included
            ExportEmptyCells = false,
            SkipEmptyRows = true,
            Indent = "  "
        };

        // Save the workbook to a memory stream using the JSON options
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, jsonOptions);
            // Convert the stream content to a UTF‑8 string
            string jsonString = Encoding.UTF8.GetString(ms.ToArray());

            // Output the JSON string (could be stored, transmitted, etc.)
            Console.WriteLine(jsonString);
        }
    }
}