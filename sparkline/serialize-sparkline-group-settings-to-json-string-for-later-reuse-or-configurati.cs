using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Json;
using System;
using System.Drawing;
using System.IO;
using System.Text;

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

        // Add a sparkline group with the data range and location
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline to the group
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Configure various sparkline group settings
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;
        group.LineWeight = 1.5;
        group.PresetStyle = SparklinePresetStyleType.Style3;
        group.PlotRightToLeft = false;

        // Set the series color
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.Orange;
        group.SeriesColor = seriesColor;

        // Prepare JSON save options (export the whole sheet)
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportArea = new CellArea { StartRow = 0, EndRow = 0, StartColumn = 0, EndColumn = 5 },
            HasHeaderRow = false,
            ExportAsString = true,
            Indent = "  "
        };

        // Serialize the workbook (including sparkline settings) to a JSON string
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, jsonOptions);
            string json = Encoding.UTF8.GetString(ms.ToArray());

            // Output the JSON string (could be stored or shared)
            Console.WriteLine(json);
        }
    }
}