// Title: Create an Excel timeline from a custom template and replace {{ProjectName}} placeholder using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook from a template file, scans all cells for the token {{ProjectName}}, substitutes it with a given project name, builds a pivot table on the Date and Project columns, attaches a timeline to the Date field, and saves the result. | Refactor the sample so that the template path, project name, and output file name are supplied via command‑line arguments while preserving the timeline generation steps. | Demonstrate how to switch the timeline’s base field from "Date" to another date column such as "StartDate" and adjust the pivot table fields accordingly with Aspose.Cells.
// Common Searches: asp.net replace placeholder in Excel template before adding timeline with Aspose.Cells | how to add a timeline to a pivot table using Aspose.Cells C# example | generate project timeline Excel file from a custom template and dynamic project name Aspose.Cells | c# Aspose.Cells create pivot table then attach timeline to date field | replace {{ProjectName}} token in Excel workbook using Aspose.Cells C#
// Tags: replace placeholder text Aspose.Cells workbook | create pivot table Aspose.Cells C# | add timeline to pivot table Aspose.Cells | load custom Excel template Aspose.Cells | dynamic project name insertion Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace TimelineFromTemplate
{
    // The example loads or creates a template workbook, replaces the {{ProjectName}} token with a specified project name, builds a pivot table on the Date and Project columns, adds a timeline linked to the Date field, sets a caption, and saves the final workbook as an Excel file.
    class Program
    {
        static void Main()
        {
            // Path to the custom template workbook
            string templatePath = "ProjectTemplate.xlsx";

            Workbook workbook = null;

            try
            {
                // Load existing template or create a new one if it does not exist
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    // Create a simple template with required columns
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Name = "Data";

                    // Header row
                    ws.Cells["A1"].PutValue("Date");
                    ws.Cells["B1"].PutValue("Project");

                    // Sample data (optional)
                    ws.Cells["A2"].PutValue(DateTime.Today);
                    ws.Cells["B2"].PutValue("Sample Project");
                }

                Worksheet sheet = workbook.Worksheets[0];

                // Dynamic project name to replace the placeholder
                string projectName = "Alpha Project";

                // Replace placeholder text "{{ProjectName}}" with the actual project name
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.Type == CellValueType.IsString && cell.StringValue.Contains("{{ProjectName}}"))
                        {
                            cell.PutValue(cell.StringValue.Replace("{{ProjectName}}", projectName));
                        }
                    }
                }

                // Determine the data range for the pivot table (including header row)
                string dataRange = $"A1:B{maxRow + 1}";

                // Add a pivot table that will serve as the data source for the timeline
                int pivotIndex = sheet.PivotTables.Add(dataRange, "D1", "ProjectPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the Date field to the Row area (base field for timeline)
                pivot.AddFieldToArea(PivotFieldType.Row, "Date");

                // Add the Project field to the Data area (just to have some data)
                pivot.AddFieldToArea(PivotFieldType.Data, "Project");

                // Refresh and calculate the pivot table
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a timeline linked to the pivot table, using the Date field as the base field
                sheet.Timelines.Add(pivot, "F1", "Date");

                // Optionally set a caption for the timeline
                Timeline timeline = sheet.Timelines[0];
                timeline.Caption = $"Timeline for {projectName}";

                // Save the resulting workbook
                workbook.Save("ProjectTimeline.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Ensure resources are released
                if (workbook != null)
                {
                    workbook.Dispose();
                }
            }
        }
    }
}
