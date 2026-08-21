// Title: Aspose.Cells .NET – AutomaticExceptTable mode, add a ListObject, and calculate a SUM formula
// Description: Create a workbook, set FormulaSettings.CalculationMode to AutomaticExceptTable, populate data, add a ListObject table, insert a SUM formula that references the table column, manually run CalculateFormula, and save the file.
// Keywords: Aspose.Cells AutomaticExceptTable | CalcModeType.AutomaticExceptTable | Add ListObject table C# | Aspose.Cells CalculateFormula | SUM formula table column | manual formula recalculation Aspose.Cells | .NET Excel table example | Workbook calculation mode Aspose
// Common Searches: Aspose.Cells set AutomaticExceptTable mode | how to add ListObject table with Aspose.Cells .NET | calculate formulas manually after adding a table Aspose | SUM formula referencing table column in Aspose.Cells | why AutomaticExceptTable requires CalculateFormula
// Developer Intent: Enable AutomaticExceptTable calculation mode, create an Excel table (ListObject), add a SUM formula that uses the table column, and trigger manual calculation to obtain the result.
// Use Cases: Generate reports where large tables are excluded from automatic recalculation until explicitly invoked. | Build financial models that add structured tables and compute totals only after data entry is complete. | Validate formula outcomes in automated tests by manually invoking CalculateFormula when AutomaticExceptTable is active.
// AI Prompts: Write C# code that sets CalcModeType.AutomaticExceptTable, adds a ListObject over a data range, inserts a SUM formula referencing the table column, and calls workbook.CalculateFormula(). | Explain why Aspose.Cells does not auto‑recalculate formulas in AutomaticExceptTable mode and how to retrieve the computed value. | Show how to access the result of a SUM formula that references a ListObject after calling CalculateFormula in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

// Create a workbook, set FormulaSettings.CalculationMode to AutomaticExceptTable, populate data, add a ListObject table, insert a SUM formula that references the table column, manually run CalculateFormula, and save the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set calculation mode to AutomaticExceptTable.
            // Excel will recalculate automatically except for tables,
            // but Aspose.Cells does not perform automatic calculation,
            // so we will invoke CalculateFormula manually.
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including header row)
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("A");
            cells["A3"].PutValue("B");
            cells["A4"].PutValue("C");
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);

            // Add an Excel table (ListObject) over the range A1:B4
            // Older Aspose.Cells versions return the index of the added table.
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SalesTable";

            // Insert a formula that sums the "Value" column of the table
            cells["C2"].Formula = "=SUM(SalesTable[Value])";

            // Manually calculate formulas (Aspose.Cells does not auto‑calculate)
            workbook.CalculateFormula();

            // Observe the recalculated result
            Console.WriteLine("Sum of Value column: " + cells["C2"].Value);

            // Save the workbook
            workbook.Save("CalculationModeDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
