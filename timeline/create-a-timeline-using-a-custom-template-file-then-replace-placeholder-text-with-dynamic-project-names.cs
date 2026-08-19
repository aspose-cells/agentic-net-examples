// Title: C# – Load a template, build a pivot table, add an Aspose.Cells timeline, and replace {{ProjectName}} placeholder
// Description: Loads Template.xlsx (or creates a new workbook), writes date/value data, creates a PivotTable, adds a Timeline linked to the Date field, scans all cells to replace the {{ProjectName}} token with a runtime project name, and saves the file as TimelineWithProjectName.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells timeline | C# add timeline to pivot table | replace placeholder in Excel C# | Excel template automation Aspose | dynamic project name Excel | pivot table timeline .NET | Aspose.Cells load template workbook
// Common Searches: How to add a timeline to a pivot table with Aspose.Cells C# | Replace {{ProjectName}} placeholder in Excel using C# | Load an existing Excel template and populate data with Aspose.Cells | Create a timeline control linked to a pivot table programmatically | Aspose.Cells replace text in all cells
// Developer Intent: Generate an Excel workbook from a predefined template, programmatically create a pivot table, attach a timeline control, and inject a project name into placeholder cells.
// Use Cases: Automated production of project schedule reports where the timeline reflects pivot data and the project name varies per run. | Reusing a single Excel template to output multiple workbooks with different project identifiers without manual editing. | Building interactive dashboards that include a timeline filter tied to pivot data for quick date range selection.
// AI Prompts: Write C# code with Aspose.Cells to load a template workbook, create a pivot table from date/value data, add a timeline bound to the Date field, replace a {{ProjectName}} placeholder with a variable, and save the result. | Show how to iterate through every cell in an Aspose.Cells worksheet to find and replace a specific placeholder string with a dynamic value. | Provide step‑by‑step instructions for adding a timeline to a pivot table in Aspose.Cells, setting its caption, and customizing its appearance.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

// Loads Template.xlsx (or creates a new workbook), writes date/value data, creates a PivotTable, adds a Timeline linked to the Date field, scans all cells to replace the {{ProjectName}} token with a runtime project name, and saves the file as TimelineWithProjectName.xlsx using Aspose.Cells for .NET.
class TimelineWithTemplate
{
    static void Main()
    {
        try
        {
            // Load the custom template workbook if it exists; otherwise create a new workbook.
            Workbook workbook;
            const string templatePath = "Template.xlsx";
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets.Add();
            }

            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Populate data that will be used for the PivotTable (Date and Value)
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");

            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
            sheet.Cells["B4"].PutValue(180);

            // -------------------------------------------------
            // 2. Create a PivotTable that will serve as the data source for the Timeline
            // -------------------------------------------------
            // Add the PivotTable to the worksheet (source range A1:B4, destination C1)
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "C1", "ProjectPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];

            // Add fields: Date as Row field, Value as Data field
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Refresh and calculate the PivotTable
            pivot.RefreshData();
            pivot.CalculateData();

            // -------------------------------------------------
            // 3. Add a Timeline linked to the PivotTable
            // -------------------------------------------------
            // Place the Timeline starting at cell E1 and bind it to the Date field
            int timelineIdx = sheet.Timelines.Add(pivot, "E1", "Date");
            Timeline timeline = sheet.Timelines[timelineIdx];
            timeline.Caption = "Project Schedule";

            // -------------------------------------------------
            // 4. Replace placeholder text with dynamic project names
            // -------------------------------------------------
            const string placeholder = "{{ProjectName}}";
            string projectName = "Apollo Expansion";

            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.Type == CellValueType.IsString && cell.StringValue.Contains(placeholder))
                    {
                        cell.PutValue(cell.StringValue.Replace(placeholder, projectName));
                    }
                }
            }

            // -------------------------------------------------
            // 5. Save the resulting workbook
            // -------------------------------------------------
            workbook.Save("TimelineWithProjectName.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
