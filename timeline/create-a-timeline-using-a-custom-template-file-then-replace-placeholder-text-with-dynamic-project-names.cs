using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file \"{templatePath}\" not found.");
                return;
            }

            // Load the custom template workbook
            Workbook workbook = new Workbook(templatePath);
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Dynamic project name to replace the placeholder
            string projectName = "Alpha Project";

            // Replace all occurrences of the placeholder {{ProjectName}} in the worksheet
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsString && cell.StringValue.Contains("{{ProjectName}}"))
                    {
                        cell.PutValue(cell.StringValue.Replace("{{ProjectName}}", projectName));
                    }
                }
            }

            // Assume the template contains date and sales data in A1:B5
            // Create a pivot table that will serve as the data source for the timeline
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table data
            pivot.RefreshData();
            pivot.CalculateData();

            // Verify that the "Date" field exists in the pivot table before adding a timeline
            bool dateFieldExists = false;
            foreach (PivotField field in pivot.RowFields)
            {
                if (field.Name.Equals("Date", StringComparison.OrdinalIgnoreCase))
                {
                    dateFieldExists = true;
                    break;
                }
            }

            if (dateFieldExists)
            {
                // Add a timeline linked to the pivot table, positioned at cell E1, using the "Date" field
                int timelineIndex = sheet.Timelines.Add(pivot, "E1", "Date");
                Timeline timeline = sheet.Timelines[timelineIndex];
                timeline.Caption = "Project Timeline";
            }
            else
            {
                Console.WriteLine("The \"Date\" field was not added to the pivot table. Timeline will not be created.");
            }

            // Save the workbook with the timeline and replaced placeholders
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to \"{resultPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}