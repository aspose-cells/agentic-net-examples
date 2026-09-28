// Title: Render a 3‑D perspective timeline from a pivot table and save it as a PNG using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, adds a pivot table, inserts a timeline linked to the Date field, applies a 45° perspective and depth via TimelineShape.ThreeDFormat, and renders the sheet to a PNG file. | Show how to adjust the width, height, and position of a TimelineShape and configure ImageOrPrintOptions for PNG output with Aspose.Cells.
// Common Searches: aspnet how to apply 3d perspective to a timeline control in Aspose.Cells | c# export worksheet containing a timeline as a PNG image using Aspose.Cells | set ThreeDFormat perspective and Z depth for TimelineShape in Aspose.Cells | render pivot table timeline to image with Aspose.Cells .NET | adjust timeline shape size and position before rendering to PNG
// Tags: Aspose.Cells timeline 3d perspective | C# ThreeDFormat timeline shape | render worksheet to PNG Aspose.Cells | pivot table linked timeline Aspose.Cells | timeline shape size position adjustment

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace Timeline3DPerspectiveDemo
{
    // Demonstrates creating a workbook, building a pivot table, adding a timeline linked to the Date field, configuring TimelineShape.ThreeDFormat for 45° perspective, depth (Z) and extrusion, resizing the timeline, and rendering the sheet to a PNG image while optionally saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate worksheet with sample data (Date and Amount)
                sheet.Cells["A1"].Value = "Date";
                sheet.Cells["B1"].Value = "Amount";
                sheet.Cells["A2"].Value = new DateTime(2023, 1, 1);
                sheet.Cells["A3"].Value = new DateTime(2023, 2, 1);
                sheet.Cells["A4"].Value = new DateTime(2023, 3, 1);
                sheet.Cells["B2"].Value = 100;
                sheet.Cells["B3"].Value = 150;
                sheet.Cells["B4"].Value = 200;

                // Create a pivot table based on the data
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "Pivot1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the Date field as a Page (filter) field – required for Timeline
                pivot.AddFieldToArea(PivotFieldType.Page, "Date");
                // Add the Amount field as Data field
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh and calculate the pivot table
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a Timeline control linked to the pivot table (Date field)
                int timelineIndex = sheet.Timelines.Add(pivot, 10, 5, "Date");
                Timeline timeline = sheet.Timelines[timelineIndex];

                // Access the underlying shape of the timeline
                TimelineShape timelineShape = timeline.Shape;

                // Apply 3‑D perspective and depth adjustments
                ThreeDFormat threeD = timelineShape.ThreeDFormat;
                threeD.Perspective = 45;      // Perspective angle (0‑120 degrees)
                threeD.Z = 30;                // Distance from the ground (depth)
                threeD.ExtrusionHeight = 15; // Optional extrusion to enhance 3‑D effect

                // Optional: adjust size and position of the timeline
                timelineShape.Width = 400;
                timelineShape.Height = 120;
                timelineShape.Top = 200;
                timelineShape.Left = 100;

                // Render the worksheet (which now contains the 3‑D timeline) to a PNG image
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    OnePagePerSheet = true
                };
                SheetRender sheetRender = new SheetRender(sheet, renderOptions);
                sheetRender.ToImage(0, "Timeline3DPerspective.png");

                // Save the workbook for reference (optional)
                workbook.Save("Timeline3DPerspective.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
