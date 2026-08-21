// Title: Add an XLOOKUP Calculated Column to an Aspose.Cells ListObject (C#) and Verify Propagation
// Description: This C# example demonstrates how to create a workbook with a lookup table, convert a range into a ListObject, resize the table to add a new column, apply an XLOOKUP formula via SetCustomCalculatedFormula, recalculate all formulas, and confirm that the formula automatically propagates to every row before saving the file.
// Keywords: Aspose.Cells | C# | ListObject | calculated column | XLOOKUP | SetCustomCalculatedFormula | table resize | structured references | formula propagation | Excel automation
// Common Searches: Aspose.Cells add XLOOKUP column to ListObject | Resize Aspose.Cells table after inserting column | SetCustomCalculatedFormula usage in C# | Automatic formula fill for ListObject rows | How to verify XLOOKUP results in Aspose.Cells
// Developer Intent: Create a ListObject column that uses XLOOKUP and ensure the formula fills all rows automatically.
// Use Cases: Generate a lookup table and retrieve matching values inside a ListObject using XLOOKUP. | Expand an existing ListObject to include a new calculated column without losing data. | Validate that a custom formula is applied consistently across every table row after calculation.
// AI Prompts: Write C# code with Aspose.Cells to add a new column to a ListObject, set an XLOOKUP formula using structured references, and recalculate the workbook. | Explain the parameters of SetCustomCalculatedFormula (isR1C1, isLocal) and how they affect formula insertion in a ListColumn. | Modify the sample to include XLOOKUP's if_not_found argument for handling missing keys.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // This C# example demonstrates how to create a workbook with a lookup table, convert a range into a ListObject, resize the table to add a new column, apply an XLOOKUP formula via SetCustomCalculatedFormula, recalculate all formulas, and confirm that the formula automatically propagates to every row before saving the file.
    public class ListObjectCalculatedColumnXLookupDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ---------- Prepare lookup table (Key -> Value) ----------
            // Headers
            cells["F1"].PutValue("LookupKey");
            cells["G1"].PutValue("LookupValue");
            // Data
            cells["F2"].PutValue("A");
            cells["G2"].PutValue(100);
            cells["F3"].PutValue("B");
            cells["G3"].PutValue(200);
            cells["F4"].PutValue("C");
            cells["G4"].PutValue(300);

            // ---------- Prepare main data that will be turned into a ListObject ----------
            // Headers
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Key");
            // Data rows
            cells["A2"].PutValue("Item1");
            cells["B2"].PutValue("A");
            cells["A3"].PutValue("Item2");
            cells["B3"].PutValue("B");
            cells["A4"].PutValue("Item3");
            cells["B4"].PutValue("C");

            // ---------- Create ListObject (table) for the main data ----------
            // Table range: A1:B4, has headers
            int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.ShowTotals = false; // not needed for this demo

            // ---------- Add a new column to the table for the XLOOKUP result ----------
            // Insert a header for the new column
            cells["C1"].PutValue("LookupResult");
            // Expand the table range to include the new column (C)
            // Resize(startRow, startColumn, totalRows, totalColumns, preserveData)
            table.Resize(0, 0, 4, 3, false);

            // Get the newly added column (index 2, zero‑based)
            ListColumn lookupColumn = table.ListColumns[2];

            // ---------- Set XLOOKUP formula for the calculated column ----------
            // Formula uses structured references: XLOOKUP([@Key], LookupKey, LookupValue)
            // The lookup range is on the same sheet (F2:G4)
            string xlookupFormula = "=XLOOKUP([@Key],F2:F4,G2:G4)";
            // Set the custom calculated formula; use A1 notation (isR1C1 = false) and local format (isLocal = false)
            lookupColumn.SetCustomCalculatedFormula(xlookupFormula, false, false);

            // ---------- Calculate all formulas ----------
            workbook.CalculateFormula();

            // ---------- Verify automatic propagation ----------
            // The formula should be applied to every data row in the column.
            for (int row = 1; row <= 3; row++) // data rows start at index 1 (row 2 in worksheet)
            {
                // Cell address for the calculated column in each row
                string cellName = $"C{row + 1}";
                Cell cell = cells[cellName];
                Console.WriteLine($"{cellName} formula: {cell.Formula}");
                Console.WriteLine($"{cellName} value  : {cell.Value}");
            }

            // ---------- Save the workbook ----------
            workbook.Save("ListObject_XLookup_CalculatedColumn.xlsx");
        }
    }
}
