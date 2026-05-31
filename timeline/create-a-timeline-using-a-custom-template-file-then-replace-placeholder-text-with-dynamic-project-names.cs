using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "ProjectTemplate.xlsx";
            const string outputPath = "ProjectTimeline.xlsx";

            // Load template if it exists; otherwise create a new workbook
            Workbook workbook = File.Exists(templatePath) ? new Workbook(templatePath) : new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Prepare data required for the Timeline (date field)
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["A2"].PutValue(DateTime.Now.AddDays(-3));
            sheet.Cells["A3"].PutValue(DateTime.Now.AddDays(-2));
            sheet.Cells["A4"].PutValue(DateTime.Now.AddDays(-1));
            sheet.Cells["A5"].PutValue(DateTime.Now);

            // Optional numeric field for the PivotTable
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // -------------------------------------------------
            // Create a PivotTable that will serve as the Timeline data source
            // -------------------------------------------------
            int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // -------------------------------------------------
            // Add the Timeline linked to the PivotTable
            // -------------------------------------------------
            // The Timeline will start at cell F1 (upper‑left corner of its range)
            sheet.Timelines.Add(pivot, "F1", "Date");

            // -------------------------------------------------
            // Replace placeholder text with dynamic project names
            // -------------------------------------------------
            string projectName = "Apollo"; // This could come from any data source

            foreach (Cell cell in sheet.Cells)
            {
                if (cell.Type == CellValueType.IsString && cell.StringValue.Contains("{{ProjectName}}"))
                {
                    string updated = cell.StringValue.Replace("{{ProjectName}}", projectName);
                    cell.PutValue(updated);
                }
            }

            // -------------------------------------------------
            // Save the resulting workbook
            // -------------------------------------------------
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}