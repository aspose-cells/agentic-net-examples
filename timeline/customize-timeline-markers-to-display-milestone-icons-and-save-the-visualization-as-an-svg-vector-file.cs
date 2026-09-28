// Title: Render an Aspose.Cells timeline with a custom milestone PNG icon to an SVG file using C#
// AI Prompts: Create a workbook, build a pivot table, add a timeline linked to the pivot, insert a PNG milestone picture aligned to the timeline shape, and export the sheet as an SVG with Aspose.Cells for .NET. | Write C# code that places a custom milestone icon on an Aspose.Cells timeline and saves the visualization as a scalable SVG vector graphic. | Show how to align a picture with a timeline shape and render the worksheet to SVG using Aspose.Cells APIs.
// Common Searches: C# Aspose.Cells attach PNG image to timeline and export as SVG | how to add a milestone icon to an Aspose.Cells timeline and save as vector graphic | export timeline with custom icon to SVG using Aspose.Cells .NET | Aspose.Cells timeline custom picture alignment example in C#
// Tags: Aspose.Cells timeline custom icon | render worksheet to SVG Aspose.Cells | add picture to timeline shape C# | pivot table linked timeline Aspose.Cells | export timeline as vector graphic .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The example demonstrates creating a workbook, populating data, generating a pivot table, adding a timeline linked to the pivot, inserting a PNG milestone picture positioned next to the timeline shape, and rendering the worksheet to an SVG file while also saving the workbook as an XLSX.
class TimelineMilestoneSvgDemo
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate sample data (date + numeric value)
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Value");

            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 5; i++)
            {
                cells[1 + i, 0].PutValue(startDate.AddMonths(i));
                cells[1 + i, 1].PutValue((i + 1) * 10);
            }

            // 3. Create a PivotTable that will be the data source of the Timeline
            int pivotIdx = sheet.PivotTables.Add("A1:B6", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            // Refresh the pivot cache and calculate data
            pivot.RefreshData();
            pivot.CalculateData();

            // 4. Add a Timeline linked to the PivotTable
            int timelineIdx = sheet.Timelines.Add(pivot, "F1", "Date");
            Timeline timeline = sheet.Timelines[timelineIdx];

            // 5. Add a milestone icon picture (if the file exists) and position it near the Timeline
            string milestoneIconPath = "milestone.png";
            if (File.Exists(milestoneIconPath))
            {
                try
                {
                    // Add picture at an arbitrary cell (row 0, column 5) – adjust as needed
                    int pictureIdx = sheet.Pictures.Add(0, 5, milestoneIconPath);
                    Picture pic = sheet.Pictures[pictureIdx];
                    pic.Width = 30;   // width in points
                    pic.Height = 30;  // height in points

                    // Align picture with the Timeline's top-left corner
                    pic.Top = timeline.Shape.Top;
                    pic.Left = timeline.Shape.Left;
                }
                catch (Exception picEx)
                {
                    Console.WriteLine($"Failed to add milestone picture: {picEx.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Milestone icon file not found: {milestoneIconPath}");
            }

            // 6. Render the worksheet (including the customized Timeline) to SVG
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                FitToViewPort = true
            };
            SheetRender renderer = new SheetRender(sheet, svgOptions);
            renderer.ToImage(0, "TimelineMilestone.svg");

            // 7. (Optional) Save the workbook for reference.
            workbook.Save("TimelineMilestoneDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
