// Title: Export an Excel Pivot Timeline to PNG with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, adds date and amount data, builds a pivot table, attaches a timeline linked to the Date field, and renders the worksheet—including the timeline with default style—to a PNG image named Timeline.png using Aspose.Cells rendering options.
// Keywords: Aspose.Cells | C# | timeline | pivot table | export PNG | render worksheet as image | Excel timeline image | default style | WorkbookRender | ImageOrPrintOptions
// Common Searches: Aspose.Cells timeline to PNG | C# export Excel timeline as image | render pivot timeline image .NET | save Excel timeline picture | how to convert timeline to PNG using Aspose.Cells
// Developer Intent: Generate a PNG image of a pivot‑table timeline from Excel data using Aspose.Cells in C#.
// Use Cases: Embed a sales timeline snapshot in a web dashboard. | Automate monthly timeline graphics for financial reports. | Create PNG timeline files for email attachments or documentation. | Produce printable timeline images for presentations.
// AI Prompts: Provide C# code that creates a pivot table, adds a linked timeline, and saves the worksheet as a PNG using Aspose.Cells. | Explain how to modify the timeline caption, colors, or style before exporting to PNG. | Show how to render only the timeline area (excluding the rest of the sheet) to a separate PNG file with Aspose.Cells. | Demonstrate error handling when exporting a timeline image in a .NET application.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Rendering;

// This C# example creates a workbook, adds date and amount data, builds a pivot table, attaches a timeline linked to the Date field, and renders the worksheet—including the timeline with default style—to a PNG image named Timeline.png using Aspose.Cells rendering options.
class TimelineToPng
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate worksheet with sample date and amount data
            cells["A1"].Value = "Date";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = new DateTime(2021, 1, 1);
            cells["B2"].Value = 100;
            cells["A3"].Value = new DateTime(2021, 2, 1);
            cells["B3"].Value = 150;
            cells["A4"].Value = new DateTime(2021, 3, 1);
            cells["B4"].Value = 200;

            // Create a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the pivot table; place it starting at cell E5 (row 4, column 4)
            int timelineIndex = sheet.Timelines.Add(pivot, 4, 4, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Sales Timeline";

            // Render the worksheet (including the timeline) to a PNG image
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                // ImageType property removed; format inferred from file extension
                OnePagePerSheet = true
            };
            WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);
            // Render the first page (index 0) to a file named "Timeline.png"
            renderer.ToImage(0, "Timeline.png");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
