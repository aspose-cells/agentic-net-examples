using System;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelineJsonDemo
{
    // DTO class that mirrors the properties we want to serialize/deserialize
    public class TimelineDto
    {
        public string Caption { get; set; }
        public string Name { get; set; }
        public int LeftPixel { get; set; }
        public int TopPixel { get; set; }
        public int WidthPixel { get; set; }
        public int HeightPixel { get; set; }
        public bool ShowHeader { get; set; }
        public bool ShowSelectionLabel { get; set; }
        public bool ShowHorizontalScrollbar { get; set; }
        public bool ShowTimeLevel { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Populate worksheet with sample data (Date and Sales)
                cells["A1"].Value = "Date";
                cells["B1"].Value = "Sales";

                cells["A2"].Value = new DateTime(2023, 1, 1);
                cells["B2"].Value = 1000;

                cells["A3"].Value = new DateTime(2023, 2, 1);
                cells["B3"].Value = 2000;

                cells["A4"].Value = new DateTime(2023, 3, 1);
                cells["B4"].Value = 3000;

                // 3. Create a pivot table that will be the data source for the timeline
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivot.RefreshData();
                pivot.CalculateData();

                // 4. Add a timeline linked to the pivot table
                // Place the timeline below the pivot (row index 5) to avoid overlapping existing data
                int timelineIdx = sheet.Timelines.Add(pivot, 5, 0, "Date");
                Timeline timeline = sheet.Timelines[timelineIdx];

                // 5. Populate the DTO with current timeline properties
                TimelineDto dto = new TimelineDto
                {
                    Caption = timeline.Caption,
                    Name = timeline.Name,
                    LeftPixel = timeline.LeftPixel,
                    TopPixel = timeline.TopPixel,
                    WidthPixel = timeline.WidthPixel,
                    HeightPixel = timeline.HeightPixel,
                    ShowHeader = timeline.ShowHeader,
                    ShowSelectionLabel = timeline.ShowSelectionLabel,
                    ShowHorizontalScrollbar = timeline.ShowHorizontalScrollbar,
                    ShowTimeLevel = timeline.ShowTimeLevel
                };

                // 6. Serialize the DTO to JSON
                string json = JsonSerializer.Serialize(dto, new JsonSerializerOptions { WriteIndented = true });
                Console.WriteLine("Serialized Timeline JSON:");
                Console.WriteLine(json);

                // 7. Modify some values in the DTO (simulating a programmatic change)
                dto.Caption = "Updated Sales Timeline";
                dto.Name = "MySalesTimeline";
                dto.LeftPixel = 150;
                dto.TopPixel = 80;
                dto.WidthPixel = 400;
                dto.HeightPixel = 120;
                dto.ShowHeader = false;
                dto.ShowSelectionLabel = true;

                // 8. Serialize the modified DTO back to JSON (optional, just to show the change)
                string modifiedJson = JsonSerializer.Serialize(dto, new JsonSerializerOptions { WriteIndented = true });
                Console.WriteLine("\nModified Timeline JSON:");
                Console.WriteLine(modifiedJson);

                // 9. Deserialize the modified JSON back into a DTO
                TimelineDto deserializedDto = JsonSerializer.Deserialize<TimelineDto>(modifiedJson);

                // 10. Apply the deserialized properties back to the Timeline object
                timeline.Caption = deserializedDto.Caption;
                timeline.Name = deserializedDto.Name;
                timeline.LeftPixel = deserializedDto.LeftPixel;
                timeline.TopPixel = deserializedDto.TopPixel;
                timeline.WidthPixel = deserializedDto.WidthPixel;
                timeline.HeightPixel = deserializedDto.HeightPixel;
                timeline.ShowHeader = deserializedDto.ShowHeader;
                timeline.ShowSelectionLabel = deserializedDto.ShowSelectionLabel;
                timeline.ShowHorizontalScrollbar = deserializedDto.ShowHorizontalScrollbar;
                timeline.ShowTimeLevel = deserializedDto.ShowTimeLevel;

                // 11. Save the workbook to verify the timeline reflects the changes
                workbook.Save("TimelineJsonDemo.xlsx");
                Console.WriteLine("\nWorkbook saved as 'TimelineJsonDemo.xlsx' with updated timeline properties.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}