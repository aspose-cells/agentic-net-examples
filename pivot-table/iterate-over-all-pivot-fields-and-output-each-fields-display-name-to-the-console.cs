using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
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
            sheet.Cells["B2"].Value = "Laptop";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Electronics";
            sheet.Cells["B3"].Value = "Phone";
            sheet.Cells["C3"].Value = 800;

            sheet.Cells["A4"].Value = "Furniture";
            sheet.Cells["B4"].Value = "Chair";
            sheet.Cells["C4"].Value = 150;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table (row, column, data)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table so that fields are initialized
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Iterate over all base fields and output each field's display name
            Console.WriteLine("Pivot Fields Display Names:");
            foreach (PivotField field in pivotTable.BaseFields)
            {
                // DisplayName defaults to the field's name unless changed explicitly
                Console.WriteLine("- " + field.DisplayName);
            }

            // Save the workbook (optional, just to complete the lifecycle)
            workbook.Save("PivotFieldsDisplayNamesDemo.xlsx");
        }
    }
}