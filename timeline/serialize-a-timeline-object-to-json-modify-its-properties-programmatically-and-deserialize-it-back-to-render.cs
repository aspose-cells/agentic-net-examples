// Title: C# – Serialize Aspose.Cells Timeline to JSON, modify it, and deserialize to update the workbook
// Description: Creates a workbook with sample data, builds a PivotTable, adds a Timeline, captures its layout in a DTO, serializes to indented JSON, edits the JSON, deserializes back, and reapplies the changes before saving the file.
// Keywords: Aspose.Cells C# timeline JSON serialization | timeline DTO System.Text.Json | modify Aspose.Cells timeline properties programmatically | deserialize timeline JSON Aspose.Cells | C# workbook timeline example | GitHub Aspose.Cells timeline demo | export timeline layout to JSON | reapply timeline settings from JSON
// Common Searches: how to serialize an Aspose.Cells timeline to JSON in C# | change timeline caption and size using JSON Aspose.Cells | deserialize timeline DTO and apply to workbook | Aspose.Cells timeline JSON example | C# code to export timeline properties to JSON
// Developer Intent: Export a Timeline's visual settings to JSON, edit the JSON, then import the changes back into the same Aspose.Cells workbook.
// Use Cases: Store timeline layout in a database as JSON for later reuse in automated report generation. | Allow non‑technical users to adjust timeline caption, position, or size via a simple JSON file. | Batch‑process multiple workbooks by applying a shared timeline configuration delivered through a web API.
// AI Prompts: Generate C# code that reads a TimelineDto JSON file and updates an existing Aspose.Cells timeline. | Explain strategies for handling missing or extra JSON properties when deserializing a TimelineDto. | Show how to serialize all timelines in a workbook into a JSON array using System.Text.Json.

using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

// Creates a workbook with sample data, builds a PivotTable, adds a Timeline, captures its layout in a DTO, serializes to indented JSON, edits the JSON, deserializes back, and reapplies the changes before saving the file.
public class TimelineDto
{
    public string Caption { get; set; }
    public string Name { get; set; }
    public int LeftPixel { get; set; }
    public int TopPixel { get; set; }
    public int WidthPixel { get; set; }
    public int HeightPixel { get; set; }
}

public class TimelineJsonDemo
{
    public static void Run()
    {
        try
        {
            // ---------- Create a new workbook ----------
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            var cells = sheet.Cells;

            // ---------- Populate worksheet with sample data ----------
            cells["A1"].Value = "Date";
            cells["B1"].Value = "Sales";

            cells["A2"].Value = new DateTime(2023, 1, 1);
            cells["B2"].Value = 1000;

            cells["A3"].Value = new DateTime(2023, 2, 1);
            cells["B3"].Value = 1500;

            cells["A4"].Value = new DateTime(2023, 3, 1);
            cells["B4"].Value = 2000;

            // ---------- Create a PivotTable that will be the data source for the Timeline ----------
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            var pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // ---------- Add a Timeline linked to the PivotTable ----------
            int timelineIdx = sheet.Timelines.Add(pivot, 0, 0, "Date");
            var timeline = sheet.Timelines[timelineIdx];

            // ---------- Capture current Timeline properties into DTO ----------
            var dto = new TimelineDto
            {
                Caption = timeline.Caption,
                Name = timeline.Name,
                LeftPixel = timeline.LeftPixel,
                TopPixel = timeline.TopPixel,
                WidthPixel = timeline.WidthPixel,
                HeightPixel = timeline.HeightPixel
            };

            // ---------- Serialize DTO to JSON ----------
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(dto, jsonOptions);
            Console.WriteLine("Serialized Timeline JSON:");
            Console.WriteLine(json);

            // ---------- Modify JSON (demo) ----------
            dto.Caption = "Updated Sales Timeline";
            dto.Name = "MyTimeline";
            dto.LeftPixel = 120;
            dto.TopPixel = 80;
            dto.WidthPixel = 400;
            dto.HeightPixel = 150;

            // ---------- Serialize modified DTO back to JSON (optional) ----------
            string modifiedJson = JsonSerializer.Serialize(dto, jsonOptions);
            Console.WriteLine("\nModified Timeline JSON:");
            Console.WriteLine(modifiedJson);

            // ---------- Deserialize JSON back to DTO ----------
            var deserializedDto = JsonSerializer.Deserialize<TimelineDto>(modifiedJson, jsonOptions);

            // ---------- Apply deserialized properties to the Timeline ----------
            if (deserializedDto != null)
            {
                timeline.Caption = deserializedDto.Caption;
                timeline.Name = deserializedDto.Name;
                timeline.LeftPixel = deserializedDto.LeftPixel;
                timeline.TopPixel = deserializedDto.TopPixel;
                timeline.WidthPixel = deserializedDto.WidthPixel;
                timeline.HeightPixel = deserializedDto.HeightPixel;
            }

            // ---------- Save the workbook ----------
            string outputPath = "TimelineJsonDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

class Program
{
    static void Main()
    {
        TimelineJsonDemo.Run();
    }
}
