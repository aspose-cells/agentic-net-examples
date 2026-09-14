// Title: How to serialize an Aspose.Cells Timeline to JSON, edit its properties in C#, and reapply the changes
// AI Prompts: Generate C# code that creates a workbook, adds a pivot table and a linked timeline, then serializes the timeline's settings into an indented JSON string. | Demonstrate deserializing the JSON into a TimelineDto, modifying fields such as Caption, Position, Size, and StartDate, and assigning the updated values back to the Aspose.Cells timeline object. | Provide a complete example that saves the workbook as an .xlsx file after applying the modified timeline configuration.
// Common Searches: C# serialize Aspose.Cells timeline to JSON and modify properties | how to change timeline start date using JSON in Aspose.Cells | example of TimelineDto for Aspose.Cells JSON deserialization | update Aspose.Cells timeline layout from JSON programmatically | save workbook after updating timeline settings in C#
// Tags: Aspose.Cells timeline JSON serialization C# | timeline property DTO for Aspose.Cells | deserialize timeline settings from JSON | programmatic timeline layout modification | save workbook after timeline update

using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineJsonDemo
{
    // DTO that mirrors the Timeline properties we want to serialize
    // Creates a workbook with sample data, adds a pivot table and a linked timeline, serializes the timeline's properties to JSON, programmatically alters selected values, deserializes the JSON back into a DTO, reapplies the changes to the timeline, and saves the workbook.
    public class TimelineDto
    {
        public string Caption { get; set; }
        public string Name { get; set; }
        public int LeftPixel { get; set; }
        public int TopPixel { get; set; }
        public int WidthPixel { get; set; }
        public int HeightPixel { get; set; }
        public bool ShowHeader { get; set; }
        public bool ShowHorizontalScrollbar { get; set; }
        public bool ShowSelectionLabel { get; set; }
        public bool ShowTimeLevel { get; set; }
        public DateTime StartDate { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                var cells = sheet.Cells;

                // Populate worksheet with sample data (date + value)
                cells["A1"].Value = "Date";
                cells["B1"].Value = "Value";

                cells["A2"].Value = DateTime.Now.AddDays(-3);
                cells["B2"].Value = 10;

                cells["A3"].Value = DateTime.Now.AddDays(-2);
                cells["B3"].Value = 20;

                cells["A4"].Value = DateTime.Now.AddDays(-1);
                cells["B4"].Value = 30;

                cells["A5"].Value = DateTime.Now;
                cells["B5"].Value = 40;

                // ---------- Create a PivotTable (data source for the timeline) ----------
                int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
                var pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");
                pivot.RefreshData();
                pivot.CalculateData();

                // ---------- Add a Timeline linked to the PivotTable ----------
                int timelineIdx = sheet.Timelines.Add(pivot, 0, 0, "Date");
                var timeline = sheet.Timelines[timelineIdx];

                // Set some initial properties
                timeline.Caption = "Initial Caption";
                timeline.Name = "InitialTimeline";
                timeline.LeftPixel = 100;
                timeline.TopPixel = 50;
                timeline.WidthPixel = 300;
                timeline.HeightPixel = 100;
                timeline.ShowHeader = true;
                timeline.ShowHorizontalScrollbar = true;
                timeline.ShowSelectionLabel = true;
                timeline.ShowTimeLevel = true;
                timeline.StartDate = DateTime.Now.AddDays(-2);

                // ---------- Serialize Timeline properties to JSON ----------
                var dto = new TimelineDto
                {
                    Caption = timeline.Caption,
                    Name = timeline.Name,
                    LeftPixel = timeline.LeftPixel,
                    TopPixel = timeline.TopPixel,
                    WidthPixel = timeline.WidthPixel,
                    HeightPixel = timeline.HeightPixel,
                    ShowHeader = timeline.ShowHeader,
                    ShowHorizontalScrollbar = timeline.ShowHorizontalScrollbar,
                    ShowSelectionLabel = timeline.ShowSelectionLabel,
                    ShowTimeLevel = timeline.ShowTimeLevel,
                    StartDate = timeline.StartDate
                };

                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(dto, jsonOptions);
                Console.WriteLine("Serialized Timeline JSON:");
                Console.WriteLine(json);

                // ---------- Modify JSON (simulated) ----------
                // For demonstration, we'll change a few values programmatically
                var modifiedDto = JsonSerializer.Deserialize<TimelineDto>(json, jsonOptions);
                modifiedDto.Caption = "Modified Caption";
                modifiedDto.Name = "ModifiedTimeline";
                modifiedDto.LeftPixel = 150;
                modifiedDto.TopPixel = 80;
                modifiedDto.WidthPixel = 350;
                modifiedDto.HeightPixel = 120;
                modifiedDto.ShowHeader = false;
                modifiedDto.StartDate = DateTime.Now.AddDays(-1);

                // ---------- Apply deserialized values back to the Timeline ----------
                timeline.Caption = modifiedDto.Caption;
                timeline.Name = modifiedDto.Name;
                timeline.LeftPixel = modifiedDto.LeftPixel;
                timeline.TopPixel = modifiedDto.TopPixel;
                timeline.WidthPixel = modifiedDto.WidthPixel;
                timeline.HeightPixel = modifiedDto.HeightPixel;
                timeline.ShowHeader = modifiedDto.ShowHeader;
                timeline.ShowHorizontalScrollbar = modifiedDto.ShowHorizontalScrollbar;
                timeline.ShowSelectionLabel = modifiedDto.ShowSelectionLabel;
                timeline.ShowTimeLevel = modifiedDto.ShowTimeLevel;
                timeline.StartDate = modifiedDto.StartDate;

                // ---------- Save the workbook ----------
                string outputPath = "TimelineJsonDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
