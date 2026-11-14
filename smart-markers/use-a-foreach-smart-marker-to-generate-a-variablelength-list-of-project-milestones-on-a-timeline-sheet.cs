// Title: Create a project milestone timeline in Excel using foreach smart markers with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells foreach smart marker example for generating a milestone timeline | How to expand rows dynamically in an Excel template using WorkbookDesigner and a List of objects | Aspose.Cells populate Excel sheet with variable‑length list of milestones in .NET | Generate project timeline Excel file from C# collection using smart markers | Using foreach smart markers to create dynamic rows in an Excel workbook with Aspose.Cells
// Tags: foreach smart marker row expansion Aspose.Cells | WorkbookDesigner data source List<Milestone> | generate dynamic Excel timeline C# | populate Excel template from collection Aspose.Cells | variable length rows Excel using smart markers

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace TimelineGenerator
{
    // Simple POCO representing a project milestone
    // The sample loads or creates an Excel template containing foreach smart markers, defines a List<Milestone> POCO, assigns it as the data source for the "Milestones" marker via WorkbookDesigner, processes the markers to expand rows for each milestone, and saves the resulting timeline workbook.
    public class Milestone
    {
        public DateTime Date { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
    }

    class Program
    {
        static void Main()
        {
            const string templatePath = "TimelineTemplate.xlsx";
            const string resultPath = "TimelineResult.xlsx";

            try
            {
                Workbook workbook;

                // Load existing template if it exists; otherwise create a minimal template programmatically
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    // Create a new workbook with a single worksheet containing smart markers
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "Timeline";

                    // Header row
                    sheet.Cells[0, 0].PutValue("Date");
                    sheet.Cells[0, 1].PutValue("Title");
                    sheet.Cells[0, 2].PutValue("Description");

                    // Row with foreach smart markers (will be expanded by WorkbookDesigner)
                    sheet.Cells[1, 0].PutValue("&=Milestones!Date");
                    sheet.Cells[1, 1].PutValue("&=Milestones!Title");
                    sheet.Cells[1, 2].PutValue("&=Milestones!Description");
                }

                // Create a WorkbookDesigner to work with smart markers
                var designer = new WorkbookDesigner(workbook);

                // Prepare a variable‑length list of milestones
                var milestones = new List<Milestone>
                {
                    new Milestone
                    {
                        Date = new DateTime(2023, 1, 15),
                        Title = "Kickoff",
                        Description = "Project kickoff meeting"
                    },
                    new Milestone
                    {
                        Date = new DateTime(2023, 2, 10),
                        Title = "Design Complete",
                        Description = "Finalize system design documents"
                    },
                    new Milestone
                    {
                        Date = new DateTime(2023, 3, 5),
                        Title = "Prototype Demo",
                        Description = "Demonstrate first prototype to stakeholders"
                    }
                    // Additional milestones can be added here
                };

                // Set the data source for the foreach smart marker named "Milestones"
                designer.SetDataSource("Milestones", milestones);

                // Process the smart markers – expands the template rows for each milestone
                designer.Process();

                // Save the generated workbook
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook generated successfully: {resultPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
