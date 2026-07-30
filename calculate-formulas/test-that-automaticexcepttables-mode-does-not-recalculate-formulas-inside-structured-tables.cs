// Title: Aspose.Cells C# – Verify AutomaticExceptTable Mode Skips Table Formula Recalculation
// Description: Shows how to enable CalcModeType.AutomaticExceptTable in Aspose.Cells for .NET, add a structured formula to a ListObject, change a source cell, recalculate, and confirm the table formula stays unchanged.
// Keywords: Aspose.Cells | AutomaticExceptTable | CalcModeType | structured table formula | ListObject | C# | .NET workbook calculation | prevent table formula update | Excel table formula mode | unit test for calculation mode
// Common Searches: Aspose.Cells AutomaticExceptTable example C# | stop table formulas from recalculating Aspose.Cells | CalcModeType AutomaticExceptTable usage | structured reference formula not updating in Aspose.Cells | how to disable automatic table calculation in .NET
// Developer Intent: Validate that setting AutomaticExceptTable disables automatic recalculation of formulas inside ListObjects.
// Use Cases: Ensure table formulas remain static after source data changes when AutomaticExceptTable is active. | Persist the calculation mode across sessions by saving and reloading the workbook. | Integrate a regression test to detect unintended changes in table formula behavior after library updates.
// AI Prompts: Create an xUnit test that verifies AutomaticExceptTable mode keeps ListObject formulas unchanged after modifying dependent cells. | Provide a concise code snippet that switches a workbook to AutomaticExceptTable, updates a source value, and demonstrates the formula result does not change. | Explain how to check that the calculation mode setting is stored in the saved .xlsx file and restored on load.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Shows how to enable CalcModeType.AutomaticExceptTable in Aspose.Cells for .NET, add a structured formula to a ListObject, change a source cell, recalculate, and confirm the table formula stays unchanged.
class AutomaticExceptTableTest
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Populate source data in column A
            cells["A1"].PutValue(1);
            cells["A2"].PutValue(2);
            cells["A3"].PutValue(3);

            // Create a table that spans A1:B3 (2 columns, 3 rows)
            int tableIndex = ws.ListObjects.Add("A1", "B3", true);
            ListObject table = ws.ListObjects[tableIndex];

            // Set column headers for clarity
            table.ListColumns[0].Name = "Value";
            table.ListColumns[1].Name = "Double";

            // Add a formula to the second column: double the value column using a structured reference
            table.PutCellFormula(0, 1, "=[@Value]*2");

            // Perform initial calculation with default mode (calculates table formulas)
            wb.CalculateFormula();

            // Capture the initial result of the table formula (cell B2 corresponds to the first data row, second column)
            double initialResult = ws.Cells["B2"].DoubleValue;
            Console.WriteLine("Initial result (B2): " + initialResult);

            // Set calculation mode to AutomaticExceptTable (formulas inside tables are NOT recalculated automatically)
            wb.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

            // Change a source value that the table formula depends on (A2)
            ws.Cells["A2"].PutValue(10); // original was 2, now 10

            // Recalculate the workbook (table formulas should stay unchanged)
            wb.CalculateFormula();

            // Capture the result after the change
            double afterResult = ws.Cells["B2"].DoubleValue;
            Console.WriteLine("After change result (B2): " + afterResult);

            // Verify that the table formula was NOT recalculated (values should be equal)
            if (Math.Abs(initialResult - afterResult) < 1e-9)
            {
                Console.WriteLine("Test passed: Table formula was not recalculated.");
            }
            else
            {
                Console.WriteLine("Test failed: Table formula was recalculated.");
            }

            // Save the workbook (optional, demonstrates that the mode is persisted)
            wb.Save("AutomaticExceptTableTest.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
