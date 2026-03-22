using System;
using Aspose.Cells;

namespace ProjectTimelineExport
{
    public class ExportTimeline
    {
        public static void Run()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define header row
            cells["A1"].Value = "Task";
            cells["B1"].Value = "Start Date";
            cells["C1"].Value = "End Date";

            // Create a date style
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "m/d/yyyy";

            // Sample hierarchical tasks (indentation via spaces)
            var tasks = new[]
            {
                new { Task = "Project Initiation", Start = new DateTime(2023, 1, 1), End = new DateTime(2023, 1, 5), Level = 0 },
                new { Task = "   Requirement Gathering", Start = new DateTime(2023, 1, 2), End = new DateTime(2023, 1, 4), Level = 1 },
                new { Task = "   Stakeholder Meeting", Start = new DateTime(2023, 1, 3), End = new DateTime(2023, 1, 3), Level = 1 },
                new { Task = "Project Planning", Start = new DateTime(2023, 1, 6), End = new DateTime(2023, 1, 10), Level = 0 },
                new { Task = "   Define Milestones", Start = new DateTime(2023, 1, 7), End = new DateTime(2023, 1, 8), Level = 1 },
                new { Task = "   Resource Allocation", Start = new DateTime(2023, 1, 9), End = new DateTime(2023, 1, 9), Level = 1 }
            };

            // Populate worksheet with task data
            int rowIndex = 1; // start after header (0‑based index)
            foreach (var t in tasks)
            {
                cells[rowIndex, 0].Value = t.Task;                     // Task name with indentation
                cells[rowIndex, 1].Value = t.Start;                    // Start date
                cells[rowIndex, 2].Value = t.End;                      // End date

                // Apply date style to the date cells
                cells[rowIndex, 1].SetStyle(dateStyle);
                cells[rowIndex, 2].SetStyle(dateStyle);

                rowIndex++;
            }

            // Auto‑fit columns for better readability
            sheet.AutoFitColumns();

            // Save the workbook as XLSX (lifecycle rule: save)
            workbook.Save("ProjectTimeline.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportTimeline.Run();
        }
    }
}