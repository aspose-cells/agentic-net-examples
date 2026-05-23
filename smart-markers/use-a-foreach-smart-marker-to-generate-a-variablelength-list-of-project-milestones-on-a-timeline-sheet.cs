using System;
using System.Collections.Generic;
using Aspose.Cells;

public class Milestone
{
    public string Name { get; set; }
    public DateTime Date { get; set; }
}

public class GenerateMilestones
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Timeline";

        // Set up column headers
        sheet.Cells["A1"].PutValue("Milestone");
        sheet.Cells["B1"].PutValue("Date");

        // Insert foreach smart markers that will repeat for each item in the data source
        // "&=Milestones.Name" repeats the Name field, "&=Milestones.Date" repeats the Date field
        sheet.Cells["A2"].PutValue("&=Milestones.Name");
        sheet.Cells["B2"].PutValue("&=Milestones.Date");

        // Prepare a variable‑length list of project milestones
        List<Milestone> milestones = new List<Milestone>
        {
            new Milestone { Name = "Kickoff", Date = new DateTime(2023, 1, 10) },
            new Milestone { Name = "Design Complete", Date = new DateTime(2023, 2, 15) },
            new Milestone { Name = "Prototype", Date = new DateTime(2023, 3, 20) },
            new Milestone { Name = "Release", Date = new DateTime(2023, 5, 5) }
        };

        // Bind the data source to the smart marker name "Milestones"
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource("Milestones", milestones);

        // Process the smart markers – rows will be generated according to the list size
        designer.Process();

        // Save the resulting workbook
        workbook.Save("ProjectMilestones.xlsx");
    }
}