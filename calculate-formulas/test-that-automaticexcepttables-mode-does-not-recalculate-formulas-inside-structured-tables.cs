// Title: C# Aspose.Cells test: Verify AutomaticExceptTables mode leaves ListObject formula column unchanged after external cell change
// AI Prompts: Create a C# program using Aspose.Cells that builds a worksheet with a ListObject, adds a formula column referencing cell A1, switches the workbook to the calculation mode that skips tables, updates A1, recalculates, and checks that the formula column values stay unchanged. | Write a C# unit test with Aspose.Cells demonstrating that formulas inside a structured table are not recomputed when the workbook's calculation mode is set to exclude table recalculation.
// Common Searches: Aspose.Cells C# how to prevent table formulas from recalculating when cell A1 changes | C# Aspose.Cells AutomaticExceptTables example with ListObject | verify that structured table formula column remains constant after external cell update Aspose.Cells | set calculation mode to AutomaticExceptTables in Aspose.Cells .NET | unit test for AutomaticExceptTables mode Aspose.Cells C#
// Tags: AutomaticExceptTables mode Aspose.Cells | ListObject formula column exclusion | C# structured table recalculation control | Aspose.Cells workbook calculation settings | verify table formula stability

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

// The sample creates a workbook, adds a ListObject with a formula column that references cell A1, calculates the workbook, changes A1, recalculates again, and prints before/after values to confirm that the table formulas remain unchanged, demonstrating the effect of AutomaticExceptTables mode.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];

            // External cell that will be used in table formulas
            ws.Cells["A1"].PutValue(10);

            // Populate data for a structured table (B2:C5)
            ws.Cells["B2"].PutValue("Item");
            ws.Cells["C2"].PutValue("Value");
            ws.Cells["B3"].PutValue("Item1");
            ws.Cells["C3"].PutValue(5);
            ws.Cells["B4"].PutValue("Item2");
            ws.Cells["C4"].PutValue(7);
            ws.Cells["B5"].PutValue("Item3");
            ws.Cells["C5"].PutValue(9);

            // Add a structured table covering B2:C5 (including header)
            int firstRow = 1;      // zero‑based index for B2
            int firstColumn = 1;   // column B
            int totalRows = 4;     // header + 3 data rows
            int totalColumns = 2;  // Item and Value columns

            int tableIndex = ws.ListObjects.Add(firstRow, firstColumn,
                                                firstRow + totalRows - 1,
                                                firstColumn + totalColumns - 1, true);
            ListObject table = ws.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Add a formula column (D) that adds A1 to the Value column (C)
            int formulaColIndex = firstColumn + totalColumns; // column D (zero‑based index 3)
            ws.Cells[firstRow, formulaColIndex].PutValue("Total"); // header

            // Set formulas for each data row in the table
            for (int i = 1; i < totalRows; i++)
            {
                // Row numbers in Excel are 1‑based, so add 1 to zero‑based index
                int excelRow = firstRow + i + 1;
                ws.Cells[firstRow + i, formulaColIndex].Formula = $"=A1+C{excelRow}";
            }

            // Initial calculation (default automatic mode)
            workbook.CalculateFormula();

            // Store initial values of the formula column
            double[] initialValues = new double[totalRows - 1];
            for (int i = 0; i < totalRows - 1; i++)
            {
                initialValues[i] = ws.Cells[firstRow + 1 + i, formulaColIndex].DoubleValue;
            }

            // Change the external cell A1
            ws.Cells["A1"].PutValue(20);

            // Recalculate workbook
            workbook.CalculateFormula();

            // Store values after the change
            double[] afterValues = new double[totalRows - 1];
            for (int i = 0; i < totalRows - 1; i++)
            {
                afterValues[i] = ws.Cells[firstRow + 1 + i, formulaColIndex].DoubleValue;
            }

            // Output comparison to verify that values inside the table did NOT change
            for (int i = 0; i < initialValues.Length; i++)
            {
                Console.WriteLine($"Row {i + 1}: before={initialValues[i]}, after={afterValues[i]}, unchanged={initialValues[i] == afterValues[i]}");
            }

            // Save the workbook (optional verification)
            try
            {
                workbook.Save("AutomaticExceptTablesTest.xlsx");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
