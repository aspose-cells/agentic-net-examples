using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Product";
        sheet.Cells["B1"].Value = "Region";
        sheet.Cells["C1"].Value = "Sales";

        sheet.Cells["A2"].Value = "Bike";
        sheet.Cells["B2"].Value = "North";
        sheet.Cells["C2"].Value = 1000;

        sheet.Cells["A3"].Value = "Bike";
        sheet.Cells["B3"].Value = "South";
        sheet.Cells["C3"].Value = 800;

        sheet.Cells["A4"].Value = "Car";
        sheet.Cells["B4"].Value = "North";
        sheet.Cells["C4"].Value = 1500;

        sheet.Cells["A5"].Value = "Car";
        sheet.Cells["B5"].Value = "South";
        sheet.Cells["C5"].Value = 1200;

        // Add a pivot table to the worksheet
        int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Refresh and calculate the pivot table data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Open a text file for logging the display names of all pivot fields
        using (StreamWriter writer = new StreamWriter("PivotFieldsDisplayNames.txt"))
        {
            // Iterate through all base fields (covers every field in the source data)
            PivotFieldCollection allFields = pivotTable.BaseFields;
            foreach (PivotField field in allFields)
            {
                // Use DisplayName if set; otherwise fall back to the field's Name
                string displayName = string.IsNullOrEmpty(field.DisplayName) ? field.Name : field.DisplayName;
                writer.WriteLine(displayName);
            }
        }

        // Save the workbook with the pivot table
        workbook.Save("PivotFieldsDemo.xlsx");
    }
}