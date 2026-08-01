// Title: Aspose.Cells for .NET – Build a Critical‑Path Timeline with Red Markers and Save as JPEG
// Description: This C# example creates an Excel workbook, adds task data with start/end dates and a critical flag, builds a pivot table, attaches a timeline linked to the start dates, colors the timeline shape red to highlight critical‑path items, and renders the sheet to a JPEG file while also saving the workbook.
// Keywords: Aspose.Cells | C# | timeline | critical path | red markers | JPEG export | pivot table | timeline shape color | render worksheet as image | project schedule visualization | Excel to JPEG
// Common Searches: Aspose.Cells timeline red color | export timeline to JPEG in C# | highlight critical tasks in Aspose.Cells timeline | create pivot table with timeline Aspose.Cells | render Excel sheet as image using Aspose.Cells | C# code for critical‑path timeline
// Developer Intent: Create a timeline that marks critical‑path tasks in red and output it as a JPEG image.
// Use Cases: Generate a project schedule Excel file, link a timeline to the start‑date pivot field, and visually emphasize critical tasks for stakeholder reports. | Automate conversion of Excel‑based timelines into high‑resolution JPEGs for dashboards or documentation pipelines. | Create a reusable routine that formats dates, builds a pivot table, applies red styling to the timeline shape, and saves both the workbook and its image.
// AI Prompts: Provide C# code using Aspose.Cells to add a timeline to a pivot table and set its shape fill and line to red. | Explain step‑by‑step how to render a worksheet containing a timeline to a JPEG image with Aspose.Cells. | Show how to format date columns and highlight critical‑path tasks in a timeline using Aspose.Cells APIs.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// This C# example creates an Excel workbook, adds task data with start/end dates and a critical flag, builds a pivot table, attaches a timeline linked to the start dates, colors the timeline shape red to highlight critical‑path items, and renders the sheet to a JPEG file while also saving the workbook.
class TimelineCriticalPathDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: Task, Start Date, End Date, Critical flag
            cells["A1"].Value = "Task";
            cells["B1"].Value = "Start";
            cells["C1"].Value = "End";
            cells["D1"].Value = "Critical";

            string[] tasks = { "Design", "Develop", "Test", "Deploy" };
            DateTime[] starts = {
                new DateTime(2023, 1, 1),
                new DateTime(2023, 1, 10),
                new DateTime(2023, 1, 20),
                new DateTime(2023, 1, 25)
            };
            DateTime[] ends = {
                new DateTime(2023, 1, 9),
                new DateTime(2023, 1, 19),
                new DateTime(2023, 1, 24),
                new DateTime(2023, 1, 30)
            };
            bool[] critical = { true, true, false, true };

            for (int i = 0; i < tasks.Length; i++)
            {
                cells[i + 1, 0].Value = tasks[i];
                cells[i + 1, 1].Value = starts[i];
                cells[i + 1, 2].Value = ends[i];
                cells[i + 1, 3].Value = critical[i] ? "Yes" : "No";
            }

            // Apply date format to the date columns
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "m/d/yyyy";
            for (int i = 1; i <= tasks.Length; i++)
            {
                cells[i, 1].SetStyle(dateStyle);
                cells[i, 2].SetStyle(dateStyle);
            }

            // Create a pivot table using the Start date as the row field
            int pivotIndex = sheet.PivotTables.Add("A1:D5", "F1", "TaskPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Start");
            pivot.AddFieldToArea(PivotFieldType.Column, "Critical");
            pivot.AddFieldToArea(PivotFieldType.Data, "Task");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the Start date field
            int timelineIndex = sheet.Timelines.Add(pivot, "H1", "Start");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Highlight critical path tasks by setting the Timeline shape fill and line to red
            TimelineShape timelineShape = timeline.Shape;
            // Use FillFormat and LineFormat to set colors
            timelineShape.FillFormat.ForeColor = Color.Red;
            timelineShape.LineFormat.ForeColor = Color.Red;

            // Render the worksheet (including the Timeline) to a JPEG image
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Jpeg,
                OnePagePerSheet = true
            };
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, "CriticalPathTimeline.jpg");

            // Save the workbook for reference
            workbook.Save("CriticalPathTimeline.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
