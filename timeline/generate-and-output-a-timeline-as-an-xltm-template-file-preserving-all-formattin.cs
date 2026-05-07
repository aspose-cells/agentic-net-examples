using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells;

// Create a new workbook and get the first worksheet
Workbook workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];
Cells cells = sheet.Cells;

// Populate sample data (including a date column with formatting)
cells[0, 0].Value = "Fruit";
cells[1, 0].Value = "Grape";
cells[2, 0].Value = "Blueberry";
cells[3, 0].Value = "Kiwi";
cells[4, 0].Value = "Cherry";

Style dateStyle = new CellsFactory().CreateStyle();
dateStyle.Custom = "m/d/yyyy";

cells[0, 1].Value = "Date";
cells[1, 1].Value = new DateTime(2021, 2, 5);
cells[2, 1].Value = new DateTime(2022, 3, 8);
cells[3, 1].Value = new DateTime(2023, 4, 10);
cells[4, 1].Value = new DateTime(2024, 5, 16);
cells[1, 1].SetStyle(dateStyle);
cells[2, 1].SetStyle(dateStyle);
cells[3, 1].SetStyle(dateStyle);
cells[4, 1].SetStyle(dateStyle);

cells[0, 2].Value = "Amount";
cells[1, 2].Value = 50;
cells[2, 2].Value = 60;
cells[3, 2].Value = 70;
cells[4, 2].Value = 80;

// Add a PivotTable based on the data range
PivotTableCollection pivots = sheet.PivotTables;
int pivotIdx = pivots.Add("=Sheet1!A1:C5", "A12", "FruitPivot");
PivotTable pivot = pivots[pivotIdx];
pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
pivot.AddFieldToArea(PivotFieldType.Column, "Date");
pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium10;
pivot.RefreshData();
pivot.CalculateData();

// Add a Timeline linked to the PivotTable (placed at A20)
TimelineCollection timelines = sheet.Timelines;
int timelineIdx = timelines.Add(pivot, "A20", "Date");
Timeline timeline = timelines[timelineIdx];

// Optional: set some visual properties via the Shape object
timeline.Shape.Title = "Sales Timeline";
timeline.Shape.Width = 400;
timeline.Shape.Height = 120;

// Save the workbook as an XLTM template preserving all data and formatting
XlsSaveOptions saveOptions = new XlsSaveOptions();
saveOptions.IsTemplate = true; // marks the file as a template (XLTM)
workbook.Save("TimelineTemplate.xlt", saveOptions);