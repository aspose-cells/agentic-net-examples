using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSubtotalDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Product";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Electronics";
            sheet.Cells["B2"].Value = "TV";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Electronics";
            sheet.Cells["B3"].Value = "Radio";
            sheet.Cells["C3"].Value = 800;

            sheet.Cells["A4"].Value = "Clothing";
            sheet.Cells["B4"].Value = "Shirt";
            sheet.Cells["C4"].Value = 500;

            sheet.Cells["A5"].Value = "Clothing";
            sheet.Cells["B5"].Value = "Pants";
            sheet.Cells["C5"].Value = 700;

            // Add a pivot table based on the data range A1:C5, place it at E3
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            // Row field: Category
            // Column field: Product
            // Data field: Sales
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // -------------------------------------------------
            // Hide subtotals for the Row field (Category)
            // -------------------------------------------------
            // Retrieve the row field object (first row field)
            PivotField rowField = pivotTable.RowFields[0];

            // Disable automatic subtotals
            rowField.IsAutoSubtotals = false;

            // Ensure that specific subtotal types are also hidden
            // (e.g., Sum, Count, Average, etc.)
            foreach (PivotFieldSubtotalType type in Enum.GetValues(typeof(PivotFieldSubtotalType)))
            {
                // Skip the 'None' type which does not represent a real subtotal
                if (type == PivotFieldSubtotalType.None) continue;
                rowField.SetSubtotals(type, false);
            }

            // -------------------------------------------------
            // (Optional) Show subtotals again for demonstration
            // -------------------------------------------------
            // To show subtotals, enable automatic subtotals and turn on a specific type
            // rowField.IsAutoSubtotals = true;
            // rowField.SetSubtotals(PivotFieldSubtotalType.Sum, true);

            // Refresh the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook in XLSX format
            workbook.Save("PivotTable_SubtotalsDemo.xlsx");
        }
    }
}