// Title: C# Aspose.Cells: Serialize, Modify, and Deserialize Timeline Settings with System.Text.Json
// Description: Demonstrates how to create a workbook, add a PivotTable, attach a Timeline, map its properties to a DTO, serialize the DTO to indented JSON, programmatically change values, deserialize the JSON, and reapply the settings to the Timeline before saving the file.
// Keywords: Aspose.Cells | Timeline | C# | System.Text.Json | serialize timeline | deserialize timeline | JSON round‑trip | pivot table | workbook automation | configuration DTO
// Common Searches: Aspose.Cells serialize timeline to JSON C# | modify timeline properties programmatically Aspose.Cells | deserialize timeline JSON and apply to workbook | timeline JSON round‑trip example Aspose.Cells | C# Aspose.Cells timeline configuration file
// Developer Intent: Export a Timeline's settings to JSON, edit them, and import the changes back into the workbook.
// Use Cases: Store and reuse Timeline layout across multiple reports via a JSON configuration file. | Allow non‑technical users to adjust Timeline captions, size, and visibility by editing a JSON file. | Synchronize Timeline appearance between a master workbook and generated dashboards through a JSON round‑trip.
// AI Prompts: Write C# code that reads a TimelineDto JSON file and applies the values to an existing Aspose.Cells Timeline object. | Extend the TimelineDto to include filter criteria and show how to serialize those additional settings. | Provide robust error‑handling patterns for deserializing Timeline JSON and updating the Timeline in Aspose.Cells.

using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineJsonDemo
{
    // DTO that mirrors the properties we want to persist for a Timeline
    // Demonstrates how to create a workbook, add a PivotTable, attach a Timeline, map its properties to a DTO, serialize the DTO to indented JSON, programmatically change values, deserialize the JSON, and reapply the settings to the Timeline before saving the file.
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
        public DateTime? StartDate { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // ---------- Create a workbook and populate sample data ----------
                Workbook workbook = new Workbook(); // create new workbook
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data: fruit, date, amount
                sheet.Cells["A1"].Value = "fruit";
                sheet.Cells["B1"].Value = "date";
                sheet.Cells["C1"].Value = "amount";

                sheet.Cells["A2"].Value = "Apple";
                sheet.Cells["B2"].Value = new DateTime(2023, 1, 5);
                sheet.Cells["C2"].Value = 120;

                sheet.Cells["A3"].Value = "Banana";
                sheet.Cells["B3"].Value = new DateTime(2023, 2, 10);
                sheet.Cells["C3"].Value = 150;

                sheet.Cells["A4"].Value = "Cherry";
                sheet.Cells["B4"].Value = new DateTime(2023, 3, 15);
                sheet.Cells["C4"].Value = 180;

                // ---------- Create a PivotTable (data source for the Timeline) ----------
                int pivotIdx = sheet.PivotTables.Add("A1:C4", "E1", "FruitPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "fruit");
                pivot.AddFieldToArea(PivotFieldType.Column, "date");
                pivot.AddFieldToArea(PivotFieldType.Data, "amount");
                pivot.RefreshData();
                pivot.CalculateData();

                // ---------- Add a Timeline linked to the PivotTable ----------
                int timelineIdx = sheet.Timelines.Add(pivot, 0, 0, "date");
                Timeline timeline = sheet.Timelines[timelineIdx];

                // Set initial properties
                timeline.Caption = "Sales Timeline";
                timeline.Name = "SalesTimeline";
                timeline.LeftPixel = 50;
                timeline.TopPixel = 30;
                timeline.WidthPixel = 400;
                timeline.HeightPixel = 120;
                timeline.ShowHeader = true;
                timeline.ShowHorizontalScrollbar = true;
                timeline.ShowSelectionLabel = true;
                timeline.ShowTimeLevel = true;
                timeline.StartDate = new DateTime(2023, 1, 1);

                // ---------- Serialize Timeline properties to JSON ----------
                TimelineDto dto = new TimelineDto
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

                // ---------- Programmatically modify JSON (simulating external change) ----------
                // For demonstration, we'll change some values directly in the DTO and re‑serialize.
                dto.Caption = "Updated Sales Timeline";
                dto.Name = "UpdatedTimeline";
                dto.LeftPixel = 100;
                dto.TopPixel = 80;
                dto.WidthPixel = 500;
                dto.HeightPixel = 150;
                dto.ShowHeader = false; // hide header
                string modifiedJson = JsonSerializer.Serialize(dto, jsonOptions);
                Console.WriteLine("\nModified Timeline JSON:");
                Console.WriteLine(modifiedJson);

                // ---------- Deserialize JSON back into a DTO ----------
                TimelineDto deserializedDto = JsonSerializer.Deserialize<TimelineDto>(modifiedJson, jsonOptions);

                // ---------- Apply deserialized values back to the Timeline object ----------
                if (deserializedDto != null)
                {
                    timeline.Caption = deserializedDto.Caption;
                    timeline.Name = deserializedDto.Name;
                    timeline.LeftPixel = deserializedDto.LeftPixel;
                    timeline.TopPixel = deserializedDto.TopPixel;
                    timeline.WidthPixel = deserializedDto.WidthPixel;
                    timeline.HeightPixel = deserializedDto.HeightPixel;
                    timeline.ShowHeader = deserializedDto.ShowHeader;
                    timeline.ShowHorizontalScrollbar = deserializedDto.ShowHorizontalScrollbar;
                    timeline.ShowSelectionLabel = deserializedDto.ShowSelectionLabel;
                    timeline.ShowTimeLevel = deserializedDto.ShowTimeLevel;
                    if (deserializedDto.StartDate.HasValue)
                        timeline.StartDate = deserializedDto.StartDate.Value;
                }

                // ---------- Save the workbook to verify the rendered Timeline ----------
                string outputPath = "TimelineJsonDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to '{Path.GetFullPath(outputPath)}'");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
