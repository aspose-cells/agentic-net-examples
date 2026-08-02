// Title: Aspose.Cells for .NET – Update ListObject column with IF formula and append rows
// Description: Demonstrates how to create a workbook, define a ListObject (table) with ID, Value, and Result columns, set the Result column to an IF expression ("High"/"Low"), insert additional rows that inherit the formula, recalculate formulas, display the results, and save the file as TableIfFormulaDemo.xlsx.
// Keywords: Aspose.Cells .NET | ListObject formula update | IF conditional column | add rows to table | recalculate formulas | C# workbook automation | Excel table programming | dynamic categorization
// Common Searches: how to set IF formula for a ListObject column in Aspose.Cells | add rows to an Aspose.Cells table and keep formulas | update ListColumn formula after table creation .NET | trigger formula calculation after inserting rows Aspose.Cells | C# Aspose.Cells conditional column example
// Developer Intent: Change the Result column of an existing ListObject to use an IF condition and programmatically add rows that automatically apply the new formula.
// Use Cases: Classify numeric entries as "High" or "Low" directly within an Excel table. | Extend a data table with new records while preserving calculated fields. | Generate reports that automatically adjust categorization thresholds via code.
// AI Prompts: Show C# code that updates a ListObject column formula to IF([@Value]>15,"High","Low") using Aspose.Cells and then adds rows that inherit the formula. | Provide an example of inserting multiple rows into an Aspose.Cells table and forcing formula recalculation. | Explain how to read back and verify the calculated values of a conditional column after modifying its formula and adding data.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableIfFormulaDemo
{
    // Demonstrates how to create a workbook, define a ListObject (table) with ID, Value, and Result columns, set the Result column to an IF expression ("High"/"Low"), insert additional rows that inherit the formula, recalculate formulas, display the results, and save the file as TableIfFormulaDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];

                // ---------- Create sample data ----------
                // Header row
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["C1"].PutValue("Result");

                // Initial data rows
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue(10);   // Expected Result: Low
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue(20);   // Expected Result: High

                // ---------- Create a table (list object) ----------
                // Table range: A1:C3 (including header)
                int tableIndex = sheet.ListObjects.Add("A1", "C3", true);
                ListObject table = sheet.ListObjects[tableIndex];

                // ---------- Update the formula of the "Result" column ----------
                // The "Result" column is the third column (index 2) in the table
                ListColumn resultColumn = table.ListColumns[2];
                // IF logic: if the Value in the current row is greater than 15, return "High", else "Low"
                resultColumn.Formula = "=IF([@Value]>15,\"High\",\"Low\")";

                // ---------- Add new rows to verify the new behavior ----------
                // Row offsets are zero‑based within the table (excluding the header row)
                int newRowOffset = table.DataRange.RowCount; // next free data row offset

                // First new row: ID = 3, Value = 5  -> Expected Result: Low
                table.PutCellValue(newRowOffset, 0, 3); // ID column
                table.PutCellValue(newRowOffset, 1, 5); // Value column

                // Second new row: ID = 4, Value = 30 -> Expected Result: High
                table.PutCellValue(newRowOffset + 1, 0, 4);
                table.PutCellValue(newRowOffset + 1, 1, 30);

                // ---------- Calculate formulas ----------
                wb.CalculateFormula();

                // ---------- Output results to console for verification ----------
                Console.WriteLine("Table contents after adding rows and calculating formulas:");
                for (int r = 0; r < table.DataRange.RowCount; r++) // iterate over data rows only
                {
                    int id = sheet.Cells[table.DataRange.FirstRow + r, table.DataRange.FirstColumn].IntValue;
                    double val = sheet.Cells[table.DataRange.FirstRow + r, table.DataRange.FirstColumn + 1].DoubleValue;
                    string res = sheet.Cells[table.DataRange.FirstRow + r, table.DataRange.FirstColumn + 2].StringValue;
                    Console.WriteLine($"ID={id}, Value={val}, Result={res}");
                }

                // Save the workbook
                wb.Save("TableIfFormulaDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
