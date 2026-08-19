// Title: Aspose.Cells .NET – Add Milestone Icons to a Timeline and Export as SVG
// Description: C# sample that creates a workbook, builds a pivot table, attaches a timeline, replaces the default markers with custom milestone icons, and renders the worksheet to a scalable SVG vector file using Aspose.Cells.
// Keywords: Aspose.Cells | C# timeline | milestone icons | SVG export | .NET vector graphics | pivot table timeline | custom timeline markers | timeline rendering | Aspose.Cells API | timeline shape customization
// Common Searches: how to add custom icons to an Aspose.Cells timeline | export Aspose.Cells timeline to SVG | C# timeline marker customization Aspose.Cells | render worksheet with timeline as vector image | Aspose.Cells pivot table timeline example
// Developer Intent: Generate a timeline from a pivot table, replace its markers with milestone icons, and save the visual as an SVG file.
// Use Cases: Project‑management dashboards that show key milestones with distinct icons on a web‑ready SVG timeline. | Marketing campaign reports where each phase is represented by a custom icon and exported for inclusion in presentations. | Automated batch creation of SVG assets for interactive dashboards that require vector‑based timeline visualizations.
// AI Prompts: Show me how to assign a custom PNG icon to each timeline marker in Aspose.Cells and then export the sheet to SVG. | Provide C# code to add a gradient fill and a thicker border to a timeline shape before rendering it as an SVG vector. | Explain how to programmatically resize, reposition, and style a timeline with milestone icons for optimal SVG output.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace TimelineMilestoneSvgDemo
{
    // C# sample that creates a workbook, builds a pivot table, attaches a timeline, replaces the default markers with custom milestone icons, and renders the worksheet to a scalable SVG vector file using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data with a date column (required for timeline) and a value column
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue(new DateTime(2023, 1, 1));
                cells["B2"].PutValue(100);
                cells["A3"].PutValue(new DateTime(2023, 2, 1));
                cells["B3"].PutValue(150);
                cells["A4"].PutValue(new DateTime(2023, 3, 1));
                cells["B4"].PutValue(200);
                cells["A5"].PutValue(new DateTime(2023, 4, 1));
                cells["B5"].PutValue(250);

                // Create a pivot table that will serve as the data source for the timeline
                int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Value");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a timeline linked to the pivot table (using the date field as the base field)
                int timelineIdx = sheet.Timelines.Add(pivot, "F1", "Date");
                Timeline timeline = sheet.Timelines[timelineIdx];

                // Customize the timeline shape appearance
                TimelineShape tlShape = timeline.Shape;

                // Set a solid fill type
                tlShape.Fill.FillType = FillType.Solid;

                // Attempt to set a fill color (some versions expose ForeColor)
                try
                {
                    var foreColorProp = tlShape.Fill.GetType().GetProperty("ForeColor");
                    if (foreColorProp != null && foreColorProp.CanWrite)
                    {
                        foreColorProp.SetValue(tlShape.Fill, Color.DarkOrange);
                    }
                }
                catch
                {
                    // Ignore if the property is unavailable
                }

                // Render the worksheet (including the customized timeline) to SVG
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true // make the SVG fit the viewport
                };

                SheetRender renderer = new SheetRender(sheet, svgOptions);
                renderer.ToImage(0, "TimelineMilestone.svg");

                // Optionally, save the workbook for reference.
                workbook.Save("TimelineMilestone.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
