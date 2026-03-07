using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;

namespace TimelineSpreadsheetMLDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (fruit, date, amount)
            cells[0, 0].Value = "fruit";
            cells[0, 1].Value = "date";
            cells[0, 2].Value = "amount";

            cells[1, 0].Value = "grape";
            cells[2, 0].Value = "blueberry";
            cells[3, 0].Value = "kiwi";
            cells[4, 0].Value = "cherry";

            // Create a date style (m/d/yyyy)
            Style dateStyle = new CellsFactory().CreateStyle();
            dateStyle.Custom = "m/d/yyyy";

            cells[1, 1].Value = new DateTime(2021, 2, 5);
            cells[2, 1].Value = new DateTime(2022, 3, 8);
            cells[3, 1].Value = new DateTime(2023, 4, 10);
            cells[4, 1].Value = new DateTime(2024, 5, 16);

            // Apply date style to the date cells
            cells[1, 1].SetStyle(dateStyle);
            cells[2, 1].SetStyle(dateStyle);
            cells[3, 1].SetStyle(dateStyle);
            cells[4, 1].SetStyle(dateStyle);

            cells[1, 2].Value = 50;
            cells[2, 2].Value = 60;
            cells[3, 2].Value = 70;
            cells[4, 2].Value = 80;

            // Add a PivotTable based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("=Sheet1!A1:C5", "A12", "DemoPivot");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "fruit");
            pivot.AddFieldToArea(PivotFieldType.Column, "date");
            pivot.AddFieldToArea(PivotFieldType.Data, "amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium10;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the PivotTable (placed at A20)
            TimelineCollection timelines = sheet.Timelines;
            int timelineIdx = timelines.Add(pivot, "A20", "date");
            Timeline timeline = timelines[timelineIdx];

            // Optional: set some Timeline properties
            timeline.Caption = "Fruit Sales Timeline";
            timeline.ShowHeader = true;
            timeline.ShowHorizontalScrollbar = true;

            // Save the workbook in SpreadsheetML (Excel 2003 XML) format
            SpreadsheetML2003SaveOptions saveOptions = new SpreadsheetML2003SaveOptions
            {
                // Include column index information for each cell (optional)
                ExportColumnIndexOfCell = true,
                // Keep XML indented for readability
                IsIndentedFormatting = true
            };
            workbook.Save("TimelineSpreadsheetML.xml", saveOptions);
        }
    }
}