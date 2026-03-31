using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace TimelineToMhtmlDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate worksheet with sample data (fruit, date, amount)
            cells[0, 0].Value = "fruit";
            cells[1, 0].Value = "grape";
            cells[2, 0].Value = "blueberry";
            cells[3, 0].Value = "kiwi";
            cells[4, 0].Value = "cherry";

            // Create a date style
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "m/d/yyyy";

            cells[0, 1].Value = "date";
            cells[1, 1].Value = new DateTime(2021, 2, 5);
            cells[2, 1].Value = new DateTime(2022, 3, 8);
            cells[3, 1].Value = new DateTime(2023, 4, 10);
            cells[4, 1].Value = new DateTime(2024, 5, 16);
            cells[1, 1].SetStyle(dateStyle);
            cells[2, 1].SetStyle(dateStyle);
            cells[3, 1].SetStyle(dateStyle);
            cells[4, 1].SetStyle(dateStyle);

            cells[0, 2].Value = "amount";
            cells[1, 2].Value = 50;
            cells[2, 2].Value = 60;
            cells[3, 2].Value = 70;
            cells[4, 2].Value = 80;

            // Add a PivotTable based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("=Sheet1!A1:C5", "A12", "FruitPivot");
            PivotTable pivot = pivots[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "fruit");
            pivot.AddFieldToArea(PivotFieldType.Column, "date");
            pivot.AddFieldToArea(PivotFieldType.Data, "amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium10;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the PivotTable's date field
            sheet.Timelines.Add(pivot, 10, 5, "date");
            Timeline timeline = sheet.Timelines[0];
            timeline.Caption = "Sales Timeline";
            timeline.Name = "SalesTimeline";
            timeline.ShowHeader = true;
            timeline.ShowHorizontalScrollbar = true;
            timeline.ShowSelectionLabel = true;
            timeline.ShowTimeLevel = true;
            timeline.StartDate = new DateTime(2021, 1, 1);
            // timeline.CurrentLevel = TimelineLevel.Month; // Removed due to unavailable enum in this version

            // Configure HTML save options for MHTML output
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);
            saveOptions.PresentationPreference = true;               // Better visual presentation
            saveOptions.ExportFrameScriptsAndProperties = true;      // Preserve interactive elements
            saveOptions.IsMobileCompatible = true;                  // Optional: mobile-friendly

            // Save the workbook as MHTML (web‑friendly) file
            string outputPath = "TimelineOutput.mht";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Timeline saved to MHTML file: {outputPath}");
        }
    }
}