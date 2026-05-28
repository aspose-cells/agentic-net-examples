using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample source data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a pivot table based on the source range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "MyPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add a row field (valid)
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");

            // *** Validation: ensure at least one data field exists before calculations ***
            if (pivot.DataFields.Count == 0)
            {
                // No data fields – handle as needed (e.g., log, throw, or skip calculation)
                Console.WriteLine("PivotTable validation failed: No data fields defined.");
            }
            else
            {
                // At least one data field exists – safe to calculate
                pivot.RefreshData();
                pivot.CalculateData();
                Console.WriteLine("PivotTable calculated successfully.");
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ValidatedPivotTable.xlsx");
        }
    }
}