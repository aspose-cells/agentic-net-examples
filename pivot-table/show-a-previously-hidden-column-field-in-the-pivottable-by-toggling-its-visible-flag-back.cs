using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotShowHiddenColumn
{
    public class ShowHiddenColumnField
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].Value = "Product";
                sheet.Cells["B1"].Value = "Region";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "Bike";
                sheet.Cells["B2"].Value = "North";
                sheet.Cells["C2"].Value = 1200;

                sheet.Cells["A3"].Value = "Bike";
                sheet.Cells["B3"].Value = "South";
                sheet.Cells["C3"].Value = 1500;

                sheet.Cells["A4"].Value = "Car";
                sheet.Cells["B4"].Value = "North";
                sheet.Cells["C4"].Value = 2000;

                sheet.Cells["A5"].Value = "Car";
                sheet.Cells["B5"].Value = "South";
                sheet.Cells["C5"].Value = 2500;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields: Product as row, Region as column, Sales as data
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Build the initial pivot
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -------------------------------------------------
                // Hide all items of the column field (Region)
                // -------------------------------------------------
                PivotField columnField = pivotTable.ColumnFields[0]; // Region field

                foreach (PivotItem item in columnField.PivotItems)
                {
                    item.IsHidden = true;
                }

                // Recalculate after hiding
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save workbook with hidden column items
                workbook.Save("PivotTable_HiddenColumnItems.xlsx");

                // -------------------------------------------------
                // Show the previously hidden column items again
                // -------------------------------------------------
                foreach (PivotItem item in columnField.PivotItems)
                {
                    item.IsHidden = false;
                }

                // Recalculate after showing
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save workbook with visible column items
                workbook.Save("PivotTable_ShownColumnItems.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }

        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}