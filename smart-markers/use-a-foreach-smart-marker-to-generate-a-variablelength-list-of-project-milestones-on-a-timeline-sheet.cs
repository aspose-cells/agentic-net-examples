// Title: Generate a dynamic project milestone timeline in Excel using foreach smart markers and WorkbookDesigner with Aspose.Cells for .NET (C#)
// Common Searches: C# foreach smart marker example to create a milestone timeline with Aspose.Cells | how to bind a List<Milestone> to smart markers using WorkbookDesigner in Aspose.Cells | populate variable‑length rows in an Excel sheet from a C# collection with Aspose.Cells | create an Excel timeline from a collection of dates and descriptions using Aspose.Cells smart markers | Aspose.Cells generate dynamic rows for project milestones using foreach smart markers
// Tags: foreach smart marker WorkbookDesigner Aspose.Cells | bind C# collection to smart markers Excel | dynamic row generation Aspose.Cells | project milestone timeline Excel C# | populate Excel sheet from List<Milestone>

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The sample creates a new workbook, adds a "Timeline" worksheet, sets header cells, and places foreach smart markers (&=Milestones.Date&, &=Milestones.Description&) that expand for each item in a C# List<Milestone>. A WorkbookDesigner binds the collection to the "Milestones" data source, processes the markers, and saves the result as Timeline.xlsx, producing a variable‑length list of project milestones.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a worksheet for the timeline
        int sheetIndex = workbook.Worksheets.Add();
        Worksheet sheet = workbook.Worksheets[sheetIndex];
        sheet.Name = "Timeline";

        // Header row
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Milestone");

        // Insert foreach smart markers (they will be expanded for each item in the collection)
        sheet.Cells["A2"].PutValue("&=Milestones.Date&");
        sheet.Cells["B2"].PutValue("&=Milestones.Description&");

        // Sample data: variable‑length list of milestones
        List<Milestone> milestones = new List<Milestone>
        {
            new Milestone { Date = new DateTime(2023, 1, 15), Description = "Project Kickoff" },
            new Milestone { Date = new DateTime(2023, 3, 10), Description = "First Release" },
            new Milestone { Date = new DateTime(2023, 6, 30), Description = "Beta Testing" },
            new Milestone { Date = new DateTime(2023, 9, 15), Description = "Final Release" }
        };

        // Process smart markers using WorkbookDesigner
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource("Milestones", milestones);
        designer.Process();

        // Save the result
        workbook.Save("Timeline.xlsx");
    }

    // Simple POCO representing a milestone
    public class Milestone
    {
        public DateTime Date { get; set; }
        public string Description { get; set; }
    }
}
