using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCalculatedFieldDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with two text columns
            cells["A1"].Value = "FirstName";
            cells["B1"].Value = "LastName";

            cells["A2"].Value = "John";
            cells["B2"].Value = "Doe";

            cells["A3"].Value = "Jane";
            cells["B3"].Value = "Smith";

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B3", "D1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Place the original text fields in the row area (optional, just for visibility)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "FirstName");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "LastName");

            // Add a calculated field that concatenates the two text fields with a hyphen
            // Formula syntax uses the '&' operator for string concatenation in Excel
            string formula = "=FirstName & \"-\" & LastName";
            pivotTable.AddCalculatedField("FullName", formula, true); // drag to data area

            // Refresh and calculate the pivot table to apply the new field
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_With_CalculatedField.xlsx");
        }
    }
}