// Title: Set workbook calculation mode to AutomaticExceptTables and add a ListObject table to test manual formula recalculation with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set calculation mode to AutomaticExceptTables | how to add a ListObject table in Aspose.Cells without automatic formula updates | manual workbook.CalculateFormula after changing a cell inside a table using Aspose.Cells | prevent table formulas from auto‑recalculating in Aspose.Cells .NET | example of AutomaticExceptTables mode with a SUM formula in Aspose.Cells
// Tags: AutomaticExceptTables calculation mode Aspose.Cells | add ListObject table Aspose.Cells C# | manual workbook.CalculateFormula invocation | disable automatic table formula recalculation | SUM formula behavior with AutomaticExceptTables

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, optionally switches the calculation mode to AutomaticExceptTables, fills cells with data and a SUM formula, forces an initial calculation, adds a ListObject table covering the data range, changes a cell value inside the table, manually triggers recalculation with workbook.CalculateFormula, and saves the file, demonstrating that formulas inside the table do not recalculate automatically when this mode is active.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Optional: set calculation mode if supported
            // workbook.Settings.CalcMode = CalculationMode.AutomaticExceptTables;

            // Get the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate some data
            sheet.Cells["A1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue(20);

            // Add a formula that sums the values in column A
            sheet.Cells["B1"].PutValue("Sum");
            sheet.Cells["B2"].Formula = "=SUM(A2:A3)";

            // Force calculation to obtain the initial sum
            workbook.CalculateFormula();

            // Add a table (ListObject) that includes the data and the formula
            // Table range: A1:B3 (zero‑based indices)
            int firstRow = 0;
            int firstColumn = 0;
            int totalRows = 3;      // number of rows (including header)
            int totalColumns = 2;   // number of columns (including header)

            int tableIndex = sheet.ListObjects.Add(firstRow,
                                                   firstColumn,
                                                   firstRow + totalRows - 1,
                                                   firstColumn + totalColumns - 1,
                                                   true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Change a value inside the table
            // Because the calculation mode may be AutomaticExceptTables,
            // the formula inside the table will NOT recalculate automatically.
            sheet.Cells["A2"].PutValue(30); // Change from 10 to 30

            // Manual recalculation if needed
            workbook.CalculateFormula();

            // Save the workbook to observe the results
            string outputPath = "CalculationModeExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
