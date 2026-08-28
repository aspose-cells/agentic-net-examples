// Title: Create a variable‑length project milestones list with foreach smart markers and add a PivotTable timeline in Aspose.Cells for .NET (C#)
// AI Prompts: Bind a List<Milestone> to foreach smart markers placed in cells A2:B2 using WorkbookDesigner to auto‑generate rows. | Construct a PivotTable from the generated milestone rows, add the Date field as a row axis and count Milestone entries as data. | Place a Timeline control linked to the PivotTable Date field at cell F1 and save the workbook as an Excel file.
// Common Searches: how to use foreach smart markers with a List<Milestone> in Aspose.Cells C# | add a timeline control to a pivot table using Aspose.Cells .NET | generate dynamic rows from a collection with WorkbookDesigner and create a timeline | Aspose.Cells example for project milestone timeline on Excel sheet | C# create pivot table and timeline from smart marker data
// Tags: foreach smart marker with WorkbookDesigner | bind collection to smart marker Aspose.Cells | create pivot table from smart marker data | add timeline control to pivot table Aspose.Cells | export project milestones to Excel C#

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace AsposeCellsSmartMarkerTimelineDemo
{
    // Simple POCO representing a project milestone
    // Demonstrates using foreach smart markers and WorkbookDesigner to expand a List<Milestone> into rows, building a PivotTable from the generated data, attaching a Timeline control to the Date field, and saving the result as an Excel workbook.
    public class Milestone
    {
        public string Name { get; set; }
        public DateTime Date { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Milestones";

                // 2. Set up header row
                sheet.Cells["A1"].PutValue("Milestone");
                sheet.Cells["B1"].PutValue("Date");

                // 3. Insert foreach smart markers in the template row (row 2)
                //    The markers will be repeated for each item in the data source collection "Milestones"
                sheet.Cells["A2"].PutValue("&=Milestones.Name");
                sheet.Cells["B2"].PutValue("&=Milestones.Date");

                // 4. Prepare a variable‑length list of milestones
                List<Milestone> milestones = new List<Milestone>
                {
                    new Milestone { Name = "Project Kick‑off", Date = new DateTime(2023, 1, 10) },
                    new Milestone { Name = "Requirement Sign‑off", Date = new DateTime(2023, 2, 5) },
                    new Milestone { Name = "Design Completion", Date = new DateTime(2023, 3, 12) },
                    new Milestone { Name = "Development Start", Date = new DateTime(2023, 4, 1) },
                    new Milestone { Name = "Testing Phase", Date = new DateTime(2023, 6, 15) },
                    new Milestone { Name = "Release", Date = new DateTime(2023, 8, 30) }
                };

                // 5. Configure WorkbookDesigner with the smart‑marker template
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Milestones", milestones);
                designer.Process(); // rows are generated automatically

                // 6. After processing, create a PivotTable based on the generated data
                int lastRow = sheet.Cells.MaxDataRow; // zero‑based index
                string sourceRange = $"A1:B{lastRow + 1}"; // Excel range is 1‑based
                int pivotIndex = sheet.PivotTables.Add(sourceRange, "D1", "MilestonePivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add "Date" as a Row field (required for Timeline)
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");
                // Add "Milestone" as a Row field to group by name
                pivot.AddFieldToArea(PivotFieldType.Row, "Milestone");
                // Add "Milestone" as a Data field to count occurrences per date
                pivot.AddFieldToArea(PivotFieldType.Data, "Milestone");
                // Optional: set a simple style
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

                // Refresh the PivotTable to calculate data
                pivot.RefreshData();
                pivot.CalculateData();

                // 7. Add a Timeline control linked to the "Date" field of the PivotTable
                //    The timeline will be placed with its upper‑left corner at cell F1
                sheet.Timelines.Add(pivot, "F1", "Date");

                // 8. Save the resulting workbook
                workbook.Save("ProjectMilestonesTimeline.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
