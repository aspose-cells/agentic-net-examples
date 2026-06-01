using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotUpdateCalculatedField
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a pivot table with a calculated field
            Workbook workbook = new Workbook("input.xlsx"); // <-- load rule

            // Assume the pivot table is on the first worksheet and named "PivotTable1"
            Worksheet sheet = workbook.Worksheets[0];
            PivotTableCollection pivotTables = sheet.PivotTables;

            // Find the pivot table by name (or use index if you know it)
            PivotTable pivotTable = null;
            foreach (PivotTable pt in pivotTables)
            {
                if (pt.Name == "PivotTable1")
                {
                    pivotTable = pt;
                    break;
                }
            }

            if (pivotTable == null)
            {
                Console.WriteLine("Pivot table 'PivotTable1' not found.");
                return;
            }

            // Name of the calculated field to be updated
            string calcFieldName = "TotalSales";

            // New formula that incorporates a discount rate parameter.
            // The discount rate can be a named range or a cell reference, e.g., a named range "DiscountRate".
            string newFormula = "=Sales*(1-DiscountRate)";

            // Add (or replace) the calculated field with the new formula.
            // Aspose.Cells will overwrite the existing calculated field if the name already exists.
            pivotTable.AddCalculatedField(calcFieldName, newFormula, true); // <-- add-calculated-field rule

            // Refresh the pivot table data and recalculate to apply the new formula
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save("output.xlsx"); // <-- save rule

            Console.WriteLine("Calculated field updated and workbook saved as 'output.xlsx'.");
        }
    }
}