// Title: C# example: Verify AutomaticExceptTables mode prevents auto‑recalculation of ListObject formulas in Aspose.Cells
// Common Searches: Aspose.Cells C# test AutomaticExceptTables mode with structured table formulas | how to ensure ListObject column formula does not recalc automatically in Aspose.Cells | sample code to check AutomaticExceptTables behavior after changing source cell in .NET | prevent automatic formula updates in Excel tables using Aspose.Cells calculation settings | manual CalculateFormula usage to refresh ListObject formulas in C#
// Tags: AutomaticExceptTables calculation mode Aspose.Cells | ListObject structured table formula C# | manual workbook.CalculateFormula Aspose.Cells | prevent auto formula recalculation Excel table | verify formula update behavior Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AutomaticExceptTablesTest
{
    // The program creates a workbook, adds a ListObject covering A1:B6, assigns a structured reference formula to the second column, runs an initial calculation, changes a source cell, and shows that the formula column does not update automatically; a subsequent workbook.CalculateFormula call refreshes the value.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Fill column A with sample data (A2:A6)
                for (int i = 0; i < 5; i++)
                {
                    sheet.Cells[i + 1, 0].PutValue(i + 1); // Values 1..5
                }

                // Define headers for the structured table
                sheet.Cells["A1"].PutValue("Data");
                sheet.Cells["B1"].PutValue("Double");

                // Add a ListObject (structured table) covering the range A1:B6
                int firstRow = 0;          // zero‑based index for row 1
                int firstColumn = 0;       // column A
                int totalRows = 6;         // rows 1‑6 (including header)
                int totalColumns = 2;      // columns A‑B
                int tableIndex = sheet.ListObjects.Add(
                    firstRow,
                    firstColumn,
                    firstRow + totalRows - 1,
                    firstColumn + totalColumns - 1,
                    true); // hasHeaders = true
                ListObject table = sheet.ListObjects[tableIndex];

                // Table visual settings
                table.ShowHeaderRow = true;
                table.ShowTableStyleFirstColumn = false;
                table.ShowTableStyleLastColumn = false;
                table.ShowTableStyleRowStripes = true;
                table.ShowTableStyleColumnStripes = false;

                // Set the formula for the second column using structured reference
                table.ListColumns[1].Formula = "=[Data]*2";

                // Perform an initial calculation so that the formula column gets values
                workbook.CalculateFormula();

                // Capture the original calculated value of the first data row (row 2, column B)
                double originalFormulaValue = sheet.Cells[1, 1].DoubleValue; // B2
                Console.WriteLine($"Original formula value (B2): {originalFormulaValue}");

                // Change the source data that the formula depends on (A2)
                sheet.Cells[1, 0].PutValue(100); // Set A2 to 100

                // Read the value after the change (automatic recalculation)
                double afterChangeFormulaValue = sheet.Cells[1, 1].DoubleValue; // B2
                Console.WriteLine($"Formula value after source change (B2) without manual calc: {afterChangeFormulaValue}");

                // Force a manual calculation to see the updated value (useful if CalcMode is Manual)
                workbook.CalculateFormula();
                double afterManualCalcValue = sheet.Cells[1, 1].DoubleValue; // B2
                Console.WriteLine($"Formula value after workbook.CalculateFormula(): {afterManualCalcValue}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
