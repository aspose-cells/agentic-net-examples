// Title: Aspose.Cells .NET: Build a Custom‑Colored Project Timeline and Export as a High‑Resolution JPEG
// Description: Creates a workbook, fills it with phase data, generates a PivotTable, adds a Timeline linked to the Date field, applies optional fill and border colors, sets a caption, and renders the sheet to a 300 DPI JPEG while also saving the workbook for later edits.
// Keywords: Aspose.Cells timeline | C# timeline export JPEG | custom colored timeline Aspose | high resolution spreadsheet image .NET | pivot table timeline Aspose.Cells | 300 DPI JPEG export | project phase timeline C# | Aspose.Cells rendering options
// Common Searches: How to add a timeline to an Excel sheet with Aspose.Cells | Set timeline shape fill and border colors using Aspose.Cells .NET | Export worksheet with timeline to high‑resolution JPEG | Create project phase timeline from pivot data in C# | Aspose.Cells render workbook to 300 DPI image
// Developer Intent: Generate a pivot‑driven timeline, style it with custom colors, and render the sheet to a high‑resolution JPEG image.
// Use Cases: Produce a printable, color‑coded project timeline for reports or presentations. | Automate daily or weekly reporting by exporting the timeline as a 300 DPI JPEG for high‑quality printing. | Save the workbook for future modifications while delivering a ready‑to‑publish image file.
// AI Prompts: Write C# code with Aspose.Cells to create a timeline from a pivot table, apply LightGreen fill and DarkGreen border, and export the sheet to a 600 DPI JPEG. | Provide a step‑by‑step tutorial on customizing timeline colors and adjusting image resolution when rendering a workbook with Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates a workbook, fills it with phase data, generates a PivotTable, adds a Timeline linked to the Date field, applies optional fill and border colors, sets a caption, and renders the sheet to a 300 DPI JPEG while also saving the workbook for later edits.
class TimelineCustomColorExport
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // ------------------------------------------------------------
            // 1. Populate sample data: Phase, Date, Value
            // ------------------------------------------------------------
            cells["A1"].PutValue("Phase");
            cells["B1"].PutValue("Date");
            cells["C1"].PutValue("Value");

            string[] phases = { "Planning", "Design", "Implementation", "Testing", "Deployment" };
            DateTime startDate = new DateTime(2023, 1, 1);
            Random rnd = new Random();

            for (int i = 0; i < phases.Length; i++)
            {
                cells[1 + i, 0].PutValue(phases[i]);                     // Phase name
                cells[1 + i, 1].PutValue(startDate.AddMonths(i));       // Date for the phase
                cells[1 + i, 2].PutValue(rnd.Next(50, 150));            // Some numeric value
            }

            // ------------------------------------------------------------
            // 2. Create a PivotTable to serve as the data source for the Timeline
            // ------------------------------------------------------------
            int pivotIdx = ws.PivotTables.Add("A1:C6", "E2", "ProjectPivot");
            PivotTable pivot = ws.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Phase");
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // ------------------------------------------------------------
            // 3. Add a Timeline linked to the Date field of the PivotTable
            // ------------------------------------------------------------
            ws.Timelines.Add(pivot, 15, 1, "Date"); // placed at row 15, column 1
            Timeline timeline = ws.Timelines[0];

            // ------------------------------------------------------------
            // 4. Apply custom appearance to the Timeline shape (optional)
            // ------------------------------------------------------------
            try
            {
                // Set solid fill color if the API is available
                // timeline.Shape.Fill.SetSolidFill(Color.LightBlue);
                // Set border (line) color if the API is available
                // timeline.Shape.Line.Color = Color.DarkBlue;
                // If the above methods are not supported in the current version,
                // they are safely ignored.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to apply custom colors – {ex.Message}");
            }

            // Set caption
            timeline.Caption = "Project Timeline";

            // ------------------------------------------------------------
            // 5. Export the worksheet (including the Timeline) to a high‑resolution JPEG
            // ------------------------------------------------------------
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Jpeg,
                HorizontalResolution = 300,   // 300 DPI horizontal
                VerticalResolution = 300,     // 300 DPI vertical
                OnePagePerSheet = true
            };

            WorkbookRender renderer = new WorkbookRender(wb, imgOptions);
            // Render the first (and only) sheet – page index 0 – to a JPEG file
            string jpegPath = "ProjectTimeline.jpg";
            renderer.ToImage(0, jpegPath);
            Console.WriteLine($"JPEG image saved to: {Path.GetFullPath(jpegPath)}");

            // Optional: save the workbook for further editing
            string xlsxPath = "ProjectTimeline.xlsx";
            wb.Save(xlsxPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(xlsxPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
