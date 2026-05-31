using System;
using System.Drawing;
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

        // Add a sparkline group (Line type) with the data range and location
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Configure some visual settings of the sparkline group
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.Orange;
        group.SeriesColor = seriesColor;
        group.ShowMarkers = true;
        group.LineWeight = 2.0;
        group.PresetStyle = SparklinePresetStyleType.Style3;

        // Prepare JSON save options
        JsonSaveOptions jsonOptions = new JsonSaveOptions();

        // Save the workbook (including sparkline settings) to a memory stream as JSON
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, jsonOptions);

            // Convert the JSON bytes to a string for reuse or sharing
            string jsonString = Encoding.UTF8.GetString(ms.ToArray());

            // Output the JSON string (could be stored, transmitted, etc.)
            Console.WriteLine(jsonString);
        }
    }
}