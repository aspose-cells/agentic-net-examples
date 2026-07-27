using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotPercentage
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Sales";
            cells["A2"].Value = "Apple";
            cells["B2"].Value = 1200;
            cells["A3"].Value = "Banana";
            cells["B3"].Value = 800;
            cells["A4"].Value = "Cherry";
            cells["B4"].Value = 500;
            cells["A5"].Value = "Date";
            cells["B5"].Value = 1500;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add row field (Product) and data field (Sales)
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Add a calculated field that simply references the Sales field.
            // The field will later be displayed as percentage of total.
            pivot.AddCalculatedField("PctOfTotal", "=Sales", true);

            // Retrieve the newly added calculated field (it is the last data field)
            PivotField pctField = pivot.DataFields[pivot.DataFields.Count - 1];

            // Configure the field to show values as percentage of total
            pctField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.PercentageOfTotal;

            // Optional: format as percentage with two decimal places
            pctField.NumberFormat = "0.00%";

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_PercentageOfTotal.xlsx");
        }
    }
}