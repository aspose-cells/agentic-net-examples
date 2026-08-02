// Title: Create an Excel Timeline from a Template and Populate Project Names – Aspose.Cells C#
// Description: Loads a template workbook (or creates a new one), adds sample project data, builds a PivotTable as the timeline source, inserts a Timeline linked to the PivotTable, uses WorkbookDesigner to replace a placeholder with a dynamic list of project names, and saves the result as TimelineWithDynamicProjects.xlsx.
// Keywords: Aspose.Cells | C# timeline | Excel timeline template | PivotTable timeline | WorkbookDesigner placeholder | dynamic project names | Excel automation | timeline slicer | template workbook | replace placeholder Aspose
// Common Searches: Aspose.Cells add timeline to pivot table C# | replace placeholder text in Excel template using Aspose.Cells | create timeline chart from template workbook .NET | WorkbookDesigner set list data source C# | generate project schedule Excel with timeline Aspose
// Developer Intent: Generate an Excel workbook that loads a template, creates a pivot‑based timeline, and injects a list of project names into placeholders using Aspose.Cells for .NET.
// Use Cases: Automated project schedule workbook where the timeline updates as new projects are added. | Reusable Excel template that can be populated with different project names at runtime. | Combined pivot data and visual timeline report for multiple projects in a single file.
// AI Prompts: Show how to load project names from a database instead of a hard‑coded list. | Explain how to style the timeline slicer and caption with Aspose.Cells. | Add robust error handling for missing template files and empty project lists.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

// Loads a template workbook (or creates a new one), adds sample project data, builds a PivotTable as the timeline source, inserts a Timeline linked to the PivotTable, uses WorkbookDesigner to replace a placeholder with a dynamic list of project names, and saves the result as TimelineWithDynamicProjects.xlsx.
class TimelineWithTemplate
{
    static void Main()
    {
        try
        {
            // Path to the template workbook.
            string templatePath = "Template.xlsx";

            // Load the template if it exists; otherwise create a new workbook.
            Workbook workbook = File.Exists(templatePath) ? new Workbook(templatePath) : new Workbook();

            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 1. Prepare sample data that will be used by the PivotTable.
            // ------------------------------------------------------------
            if (cells.MaxDataRow < 4) // simple check to avoid duplicate data
            {
                // Header row
                cells["A1"].PutValue("Project");
                cells["B1"].PutValue("Date");
                cells["C1"].PutValue("Value");

                // Sample rows
                cells["A2"].PutValue("Alpha");
                cells["B2"].PutValue(new DateTime(2023, 1, 1));
                cells["C2"].PutValue(120);

                cells["A3"].PutValue("Beta");
                cells["B3"].PutValue(new DateTime(2023, 2, 15));
                cells["C3"].PutValue(200);

                cells["A4"].PutValue("Gamma");
                cells["B4"].PutValue(new DateTime(2023, 3, 10));
                cells["C4"].PutValue(150);
            }

            // ------------------------------------------------------------
            // 2. Create a PivotTable that will serve as the data source for the Timeline.
            // ------------------------------------------------------------
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "C6", "ProjectPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the PivotTable fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Project");   // Row field
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");   // Column field (time axis)
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");    // Data field

            // Refresh the PivotTable so that it contains the latest data
            pivot.RefreshData();
            pivot.CalculateData();

            // ------------------------------------------------------------
            // 3. Add a Timeline linked to the PivotTable.
            // ------------------------------------------------------------
            int timelineIndex = sheet.Timelines.Add(pivot, "E1", "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Project Timeline";

            // ------------------------------------------------------------
            // 4. Replace placeholder text with dynamic project names using WorkbookDesigner.
            // ------------------------------------------------------------
            List<string> projectNames = new List<string> { "Alpha", "Beta", "Gamma" };
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("ProjectName", projectNames);
            designer.Process();

            // ------------------------------------------------------------
            // 5. Save the resulting workbook.
            // ------------------------------------------------------------
            string outputPath = "TimelineWithDynamicProjects.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
