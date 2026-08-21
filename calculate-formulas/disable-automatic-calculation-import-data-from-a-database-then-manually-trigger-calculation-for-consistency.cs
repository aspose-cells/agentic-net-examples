// Title: Disable Auto‑Calc, Import DataTable, and Manually Recalculate Formulas with Aspose.Cells for .NET
// Description: Demonstrates how to set Aspose.Cells calculation mode to Manual, import a DataTable (simulating database rows) into a worksheet with headers, add SUM and SUMPRODUCT formulas, trigger a single Workbook.CalculateFormula call, and save the workbook as XLSX.
// Keywords: Aspose.Cells manual calculation | CalcModeType.Manual | Workbook.CalculateFormula | ImportData DataTable | ImportTableOptions | C# Excel automation | disable automatic formula evaluation | bulk data import Excel | calculate formulas after import
// Common Searches: Aspose.Cells turn off automatic calculation .NET | Import DataTable into Excel worksheet using Aspose.Cells | Manual formula evaluation Aspose.Cells C# | CalcModeType.Manual example | Workbook.CalculateFormula after data import
// Developer Intent: The developer needs to prevent formulas from recalculating while loading data, then evaluate all formulas in one explicit step.
// Use Cases: Load thousands of rows from a database into a template without triggering per‑row recalculation, then compute totals once. | Create financial or inventory reports where data is staged first and formulas are applied only after the dataset is complete. | Build an Excel export service that inserts external data, keeps formulas dormant, and activates them with a single calculate call before delivering the file.
// AI Prompts: Show C# code that disables automatic calculation in Aspose.Cells, imports a DataTable with column headers, adds dependent formulas, and calls Workbook.CalculateFormula. | Provide an Aspose.Cells example using ImportTableOptions (IsFieldNameShown, InsertRows) while the workbook is in manual calculation mode. | Explain best practices for bulk importing database rows into an Excel file with Aspose.Cells and performing a one‑time formula evaluation.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsCalcExample
{
    // Demonstrates how to set Aspose.Cells calculation mode to Manual, import a DataTable (simulating database rows) into a worksheet with headers, add SUM and SUMPRODUCT formulas, trigger a single Workbook.CalculateFormula call, and save the workbook as XLSX.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ----- Disable automatic calculation -----
            // Set calculation mode to Manual so formulas are not evaluated automatically
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
            // Optional: ensure formulas are not calculated on open
            workbook.Settings.FormulaSettings.CalculateOnOpen = false;

            // ----- Simulate importing data from a database -----
            // In a real scenario you would use a SqlDataReader or similar.
            // Here we create a DataTable to represent the data source.
            DataTable dt = new DataTable("Products");
            dt.Columns.Add("ProductID", typeof(int));
            dt.Columns.Add("ProductName", typeof(string));
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("UnitPrice", typeof(double));

            dt.Rows.Add(1, "Apple", 50, 0.5);
            dt.Rows.Add(2, "Banana", 30, 0.3);
            dt.Rows.Add(3, "Cherry", 20, 1.2);

            // Import the DataTable starting at cell A1 (row 0, column 0)
            // Use ImportTableOptions to include column headers
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,
                InsertRows = true
            };
            cells.ImportData(dt, 0, 0, importOptions);

            // ----- Add sample formulas that depend on the imported data -----
            // Total quantity (sum of Quantity column)
            cells["E2"].Formula = "=SUM(C2:C4)";
            // Total value (Quantity * UnitPrice)
            cells["F2"].Formula = "=SUMPRODUCT(C2:C4, D2:D4)";

            // ----- Manually trigger calculation -----
            // Since calculation mode is Manual, we need to call CalculateFormula explicitly.
            workbook.CalculateFormula();

            // ----- Save the workbook -----
            workbook.Save("ManualCalc_ImportedData.xlsx", SaveFormat.Xlsx);
        }
    }
}
