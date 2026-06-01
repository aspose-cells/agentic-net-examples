using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Simulate data import from a database using a DataTable
            DataTable dbTable = new DataTable("SalesData");
            dbTable.Columns.Add("Product", typeof(string));
            dbTable.Columns.Add("Quantity", typeof(int));
            dbTable.Columns.Add("UnitPrice", typeof(double));

            dbTable.Rows.Add("Apple", 10, 0.5);
            dbTable.Rows.Add("Banana", 20, 0.3);
            dbTable.Rows.Add("Cherry", 15, 0.8);

            // Set import options (show column headers)
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true
            };

            // Import the data starting at cell A1 (lifecycle rule: import data)
            cells.ImportData(dbTable, 0, 0, importOptions);

            // Add a formula that calculates total sales per row
            // Assuming data starts at row 2 (index 1) after headers
            for (int row = 1; row <= dbTable.Rows.Count; row++)
            {
                // Column D will hold Quantity * UnitPrice
                cells[row, 3].Formula = $"=B{row + 1}*C{row + 1}";
            }

            // Add a summary formula to sum total sales
            int summaryRow = dbTable.Rows.Count + 2; // one extra row after data
            cells[summaryRow, 2].PutValue("Total Sales:");
            cells[summaryRow, 3].Formula = $"=SUM(D2:D{summaryRow})";

            // Disable automatic calculation
            FormulaSettings formulaSettings = workbook.Settings.FormulaSettings;
            formulaSettings.CalculationMode = CalcModeType.Manual;
            formulaSettings.CalculateOnOpen = false;
            formulaSettings.CalculateOnSave = false;

            // Manually trigger calculation to ensure consistency
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ImportedDataWithManualCalc.xlsx");
        }
    }
}