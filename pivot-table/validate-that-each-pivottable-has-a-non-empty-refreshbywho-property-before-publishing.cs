using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableValidationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Drink");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Food");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("Drink");
            sheet.Cells["B5"].PutValue(70);

            // Add a pivot table
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Amount

            // Refresh the pivot table so that RefreshedByWho gets a value
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Validate each PivotTable has a non‑empty RefreshedByWho property
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (PivotTable pt in ws.PivotTables)
                {
                    string refreshedBy = pt.RefreshedByWho;
                    if (string.IsNullOrWhiteSpace(refreshedBy))
                    {
                        throw new InvalidOperationException(
                            $"PivotTable \"{pt.Name}\" in worksheet \"{ws.Name}\" does not have a valid RefreshedByWho value.");
                    }
                }
            }

            // All validations passed – save the workbook (publishing)
            workbook.Save("ValidatedPivotTable.xlsx");
        }
    }
}