// Title: Add a Milestone Icon to an Aspose.Cells Timeline and Export the Sheet as SVG (C#)
// Description: Creates a workbook with date‑sales data, builds a pivot table, attaches a timeline, inserts a custom milestone PNG as a free‑floating shape centered on the timeline, and renders the worksheet to a scalable SVG file (optionally saving the XLSX).
// Keywords: Aspose.Cells timeline icon | C# timeline milestone | export timeline to SVG | add picture to Aspose.Cells timeline | render worksheet as SVG | free floating shape Aspose.Cells | .NET spreadsheet visualization
// Common Searches: Aspose.Cells add custom icon to timeline | Export timeline chart to SVG with Aspose.Cells | Position picture on timeline shape C# | How to render Aspose.Cells worksheet as SVG | Milestone marker on Aspose.Cells timeline
// Developer Intent: Insert a custom milestone image onto a timeline linked to a pivot table and save the visual as an SVG vector graphic.
// Use Cases: Build interactive sales dashboards where each key date is highlighted with a bespoke icon and publish the view as a resolution‑independent SVG for web embedding. | Generate printable vector reports that combine pivot‑driven timelines with branded milestone symbols, ensuring crisp scaling on any device. | Automate workbook creation that adds data, creates a timeline with image markers, and outputs both XLSX and SVG files for downstream analytics or documentation.
// AI Prompts: Generate C# code using Aspose.Cells to place a PNG milestone icon on a timeline and align it vertically in the middle of the timeline shape. | Show how to render a worksheet containing a timeline to an SVG file with FitToViewPort enabled in Aspose.Cells. | Provide a robust file‑existence check for a milestone image before inserting it into a timeline, including fallback messaging.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTimelineSvgDemo
{
    // Creates a workbook with date‑sales data, builds a pivot table, attaches a timeline, inserts a custom milestone PNG as a free‑floating shape centered on the timeline, and renders the worksheet to a scalable SVG file (optionally saving the XLSX).
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Populate sample data (date column is required for a timeline)
                cells["A1"].PutValue("Date");
                cells["A2"].PutValue(new DateTime(2023, 1, 1));
                cells["A3"].PutValue(new DateTime(2023, 2, 1));
                cells["A4"].PutValue(new DateTime(2023, 3, 1));
                cells["A5"].PutValue(new DateTime(2023, 4, 1));

                cells["B1"].PutValue("Sales");
                cells["B2"].PutValue(120);
                cells["B3"].PutValue(150);
                cells["B4"].PutValue(180);
                cells["B5"].PutValue(200);

                // 3. Create a pivot table that will serve as the timeline data source
                int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivot.RefreshData();
                pivot.CalculateData();

                // 4. Add a timeline linked to the pivot table (starts at cell F1)
                int timelineIdx = sheet.Timelines.Add(pivot, 0, 5, "Date");
                Timeline timeline = sheet.Timelines[timelineIdx];

                // 5. Add a milestone icon picture if the file exists
                string milestoneIconPath = "milestone.png";
                if (File.Exists(milestoneIconPath))
                {
                    int pictureIdx = sheet.Pictures.Add(0, 5, milestoneIconPath);
                    Picture milestonePic = sheet.Pictures[pictureIdx];

                    // Make the picture free‑floating
                    milestonePic.Placement = PlacementType.FreeFloating;

                    // Resize the icon
                    milestonePic.Width = 30;   // pixels
                    milestonePic.Height = 30; // pixels

                    // Vertically center the picture on the timeline
                    double timelineTop = timeline.Shape.Top;
                    double timelineHeight = timeline.Shape.Height;
                    milestonePic.Top = (int)(timelineTop + (timelineHeight - milestonePic.Height) / 2);
                }
                else
                {
                    Console.WriteLine($"Warning: '{milestoneIconPath}' not found. Skipping picture insertion.");
                }

                // 6. Render the worksheet (which now contains the timeline and optional icon) to SVG
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true // make the SVG fit the viewport
                };

                SheetRender renderer = new SheetRender(sheet, svgOptions);
                renderer.ToImage(0, "TimelineWithMilestone.svg");

                // 7. Save the workbook (optional, for verification)
                workbook.Save("TimelineWithMilestone.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
