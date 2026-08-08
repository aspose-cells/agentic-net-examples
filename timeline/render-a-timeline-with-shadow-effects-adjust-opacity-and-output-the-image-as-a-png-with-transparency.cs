// Title: Render a PivotTable Timeline with Shadow and Transparent PNG in C# using Aspose.Cells
// Description: Creates a workbook, builds a PivotTable, adds a linked Timeline, applies a semi‑transparent shadow, and exports the Timeline shape as a PNG with a transparent background via Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | Timeline | ShadowEffect | Transparent PNG | PivotTable timeline export | ImageOrPrintOptions | timeline rendering | PNG transparency | Aspose.Cells .NET
// Common Searches: Aspose.Cells timeline shadow effect C# | export timeline shape to PNG with transparency | how to add shadow to Aspose.Cells timeline | render timeline as image Aspose.Cells .NET | transparent background PNG Aspose.Cells timeline
// Developer Intent: Produce a PNG image of a PivotTable timeline that includes a configurable shadow and a transparent background using Aspose.Cells for .NET.
// Use Cases: Embedding timeline graphics with depth effects into web dashboards | Generating transparent PNGs for presentations that need to overlay on varied backgrounds | Automating batch creation of styled timeline images for periodic sales reports
// AI Prompts: Write C# code to apply a 40% transparent shadow to an Aspose.Cells timeline and save it as a PNG with a transparent background. | Explain how to modify shadow angle, distance, blur, and size for a timeline shape in Aspose.Cells. | Show how to loop through multiple timelines in a workbook and export each to a separate transparent PNG file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace TimelineRenderingDemo
{
    // Creates a workbook, builds a PivotTable, adds a linked Timeline, applies a semi‑transparent shadow, and exports the Timeline shape as a PNG with a transparent background via Aspose.Cells for .NET.
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

                // Populate sample data with a date column (required for Timeline)
                cells[0, 0].Value = "Product";
                cells[0, 1].Value = "Date";
                cells[0, 2].Value = "Sales";

                cells[1, 0].Value = "Apple";
                cells[1, 1].Value = new DateTime(2022, 1, 15);
                cells[1, 2].Value = 1200;

                cells[2, 0].Value = "Banana";
                cells[2, 1].Value = new DateTime(2022, 2, 10);
                cells[2, 2].Value = 950;

                cells[3, 0].Value = "Cherry";
                cells[3, 1].Value = new DateTime(2022, 3, 5);
                cells[3, 2].Value = 780;

                // Create a PivotTable based on the data
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIdx = pivots.Add("=Sheet1!A1:C4", "E5", "SalesPivot");
                PivotTable pivot = pivots[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Product");
                pivot.AddFieldToArea(PivotFieldType.Column, "Date");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
                // Add the date field to the Page (filter) area – required for Timeline
                pivot.AddFieldToArea(PivotFieldType.Page, "Date");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a Timeline linked to the PivotTable's date field
                sheet.Timelines.Add(pivot, 15, 0, "Date");
                Timeline timeline = sheet.Timelines[0];

                // Access the underlying shape of the Timeline
                TimelineShape timelineShape = timeline.Shape;

                // Configure shadow effect with desired transparency (e.g., 40% transparent)
                ShadowEffect shadow = timelineShape.ShadowEffect;
                shadow.Transparency = 0.4; // 0.0 = opaque, 1.0 = fully transparent
                shadow.Angle = 135;
                shadow.Blur = 20;
                shadow.Size = 1.0;
                shadow.Distance = 10;

                // Prepare image options: PNG format with transparent background
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    Transparent = true // Enable transparent background
                };

                // Render the Timeline shape to a PNG file
                string outputImagePath = "TimelineWithShadow.png";
                timelineShape.ToImage(outputImagePath, imgOptions);

                // Optionally save the workbook for reference
                string workbookPath = "TimelineDemo.xlsx";
                workbook.Save(workbookPath);

                Console.WriteLine($"Timeline rendered to image: {Path.GetFullPath(outputImagePath)}");
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
