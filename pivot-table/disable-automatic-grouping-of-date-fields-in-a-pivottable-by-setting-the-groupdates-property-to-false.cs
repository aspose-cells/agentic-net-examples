using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class DisableAutoDateGrouping
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a Date column and a Sales column
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue(new DateTime(2023, 1, 2));
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["A4"].PutValue(new DateTime(2023, 1, 3));
            sheet.Cells["B4"].PutValue(1800);
            sheet.Cells["A5"].PutValue(new DateTime(2023, 1, 4));
            sheet.Cells["B5"].PutValue(2000);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the Date field as a row field and Sales as a data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Aspose.Cells does not automatically group date fields; no explicit property is required.
            // If needed, grouping can be performed via the Group method; here we keep dates ungrouped.

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            string outputPath = "PivotTable_NoAutoDateGrouping.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}