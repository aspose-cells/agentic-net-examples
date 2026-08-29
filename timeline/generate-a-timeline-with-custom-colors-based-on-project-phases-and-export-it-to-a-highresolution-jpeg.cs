// Title: Create a project‑phase timeline from a PivotTable and export it as a high‑resolution JPEG with Aspose.Cells for .NET
// AI Prompts: Build a PivotTable from phase start dates, attach a Timeline control, set a custom caption, and render the sheet to a 300 dpi JPEG using Aspose.Cells. | Configure ImageOrPrintOptions for JPEG output at 300 dpi and generate the image file from the first worksheet. | Save the workbook after adding the Timeline so it can be edited later, then export only the visual timeline as a high‑resolution image.
// Common Searches: Aspose.Cells how to render a worksheet to a 300 dpi JPEG image | add a timeline linked to a pivot table date field using Aspose.Cells .NET | export Excel timeline as high resolution JPEG with Aspose.Cells | set custom caption for timeline in Aspose.Cells workbook | save workbook after creating timeline for further editing Aspose.Cells
// Tags: pivot table to timeline Aspose.Cells | render worksheet to high DPI JPEG Aspose.Cells | custom timeline caption Aspose.Cells | export Excel as 300 dpi image .NET | project phase timeline automation Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Demonstrates creating sample project phase data, building a PivotTable, adding a linked Timeline with a custom caption, and rendering the first worksheet to a 300 dpi JPEG while also saving the workbook for later edits.
class TimelineExportDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 1. Populate sample data for project phases
            // -------------------------------------------------
            cells["A1"].PutValue("Phase");
            cells["B1"].PutValue("Start");
            cells["C1"].PutValue("End");
            cells["D1"].PutValue("Progress");

            cells["A2"].PutValue("Planning");
            cells["B2"].PutValue(new DateTime(2023, 1, 1));
            cells["C2"].PutValue(new DateTime(2023, 2, 15));
            cells["D2"].PutValue(100);

            cells["A3"].PutValue("Execution");
            cells["B3"].PutValue(new DateTime(2023, 2, 16));
            cells["C3"].PutValue(new DateTime(2023, 6, 30));
            cells["D3"].PutValue(80);

            cells["A4"].PutValue("Closure");
            cells["B4"].PutValue(new DateTime(2023, 7, 1));
            cells["C4"].PutValue(new DateTime(2023, 7, 31));
            cells["D4"].PutValue(50);

            // -------------------------------------------------
            // 2. Create a PivotTable to serve as the Timeline source
            // -------------------------------------------------
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("A1:D4", "F1", "ProjectPivot");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Start");
            pivot.AddFieldToArea(PivotFieldType.Column, "Phase");
            pivot.AddFieldToArea(PivotFieldType.Data, "Progress");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

            // Refresh and calculate pivot data
            pivot.RefreshData();
            pivot.CalculateData();

            // -------------------------------------------------
            // 3. Add a Timeline linked to the "Start" date field (if supported)
            // -------------------------------------------------
            int timelineIdx = sheet.Timelines.Add(pivot, 10, 1, "Start");
            sheet.Timelines[timelineIdx].Caption = "Project Phases Timeline";

            // -------------------------------------------------
            // 4. Render the workbook (first sheet) to a high‑resolution JPEG
            // -------------------------------------------------
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Jpeg,
                HorizontalResolution = 300,   // high DPI
                VerticalResolution = 300,
                OnePagePerSheet = true
            };

            WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);
            // Save the rendered image; page index 0 corresponds to the first worksheet
            renderer.ToImage(0, "ProjectTimeline.jpg");

            // -------------------------------------------------
            // 5. Optionally save the workbook for further editing
            // -------------------------------------------------
            workbook.Save("ProjectTimeline.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
