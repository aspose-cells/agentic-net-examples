// Title: C# – Add an XLOOKUP Calculated Column to a ListObject and Verify Propagation with Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a workbook with a lookup table and a main ListObject, then adds a calculated column using SetCustomCalculatedFormula and an XLOOKUP formula. The code calculates the workbook, prints each cell's formula to confirm automatic propagation, and saves the file as CalculatedColumn_XLookup.xlsx.
// Keywords: Aspose.Cells | C# | .NET | ListObject | calculated column | XLOOKUP | SetCustomCalculatedFormula | table formula propagation | lookup table example | Excel automation
// Common Searches: Aspose.Cells add XLOOKUP column to ListObject | C# set custom calculated formula in table | verify calculated column propagation Aspose.Cells | how to use XLOOKUP with ListColumn in .NET | Aspose.Cells example for lookup table and XLOOKUP
// Developer Intent: Create a ListObject with a calculated column that uses XLOOKUP and ensure the formula is automatically applied to every data row.
// Use Cases: Generate a lookup table and retrieve matching values in a main table via XLOOKUP. | Provide a default result for missing keys using the not_found argument of XLOOKUP. | Automatically recalculate the workbook and validate that each cell in the calculated column contains the correct formula.
// AI Prompts: Write C# code with Aspose.Cells that adds a calculated column using XLOOKUP to an existing ListObject and displays each cell's formula. | Show how to create two tables, set a custom calculated formula on a ListColumn with XLOOKUP, and confirm formula propagation across rows. | Explain how to change the XLOOKUP not_found value in the calculated column and update the workbook accordingly.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // This Aspose.Cells for .NET example creates a workbook with a lookup table and a main ListObject, then adds a calculated column using SetCustomCalculatedFormula and an XLOOKUP formula. The code calculates the workbook, prints each cell's formula to confirm automatic propagation, and saves the file as CalculatedColumn_XLookup.xlsx.
    public class ListObjectCalculatedColumnWithXLookup
    {
        public static void Main()
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Create Lookup Table ----------
            // Headers
            sheet.Cells["D1"].PutValue("Key");
            sheet.Cells["E1"].PutValue("Value");
            // Data
            sheet.Cells["D2"].PutValue("A");
            sheet.Cells["E2"].PutValue(100);
            sheet.Cells["D3"].PutValue("B");
            sheet.Cells["E3"].PutValue(200);
            sheet.Cells["D4"].PutValue("C");
            sheet.Cells["E4"].PutValue(300);

            // Add the lookup ListObject (table)
            int lookupTableIdx = sheet.ListObjects.Add("D1", "E4", true);
            ListObject lookupTable = sheet.ListObjects[lookupTableIdx];
            // Optional: set a display name for the table (used in formulas)
            lookupTable.DisplayName = "LookupTable";

            // ---------- Create Main Table ----------
            // Headers (four columns, the last will be the calculated column)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Key");
            sheet.Cells["C1"].PutValue("Amount");
            sheet.Cells["D1"].PutValue("LookupValue"); // Calculated column header

            // Data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("A");
            sheet.Cells["C2"].PutValue(10);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("B");
            sheet.Cells["C3"].PutValue(20);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("D"); // This key does not exist in lookup table
            sheet.Cells["C4"].PutValue(30);

            // Add the main ListObject (table)
            int mainTableIdx = sheet.ListObjects.Add("A1", "D4", true);
            ListObject mainTable = sheet.ListObjects[mainTableIdx];
            mainTable.DisplayName = "MainTable";

            // ---------- Add Calculated Column using XLOOKUP ----------
            // The calculated column is the fourth column (index 3)
            ListColumn calcColumn = mainTable.ListColumns[3];

            // XLOOKUP searches the Key in the lookup table and returns the corresponding Value,
            // or "NotFound" if the key is absent.
            string xlookupFormula = "=XLOOKUP([@Key], LookupTable[Key], LookupTable[Value], \"NotFound\")";
            calcColumn.SetCustomCalculatedFormula(xlookupFormula, false, false);

            // ---------- Calculate and Verify Propagation ----------
            workbook.CalculateFormula();

            // Output the formulas applied to each data cell in the calculated column
            Console.WriteLine("Formulas in the calculated column:");
            int startRow = mainTable.DataRange.FirstRow;      // first data row (zero‑based)
            int startCol = mainTable.DataRange.FirstColumn;  // first data column (zero‑based)
            int dataRowCount = mainTable.DataRange.RowCount; // number of data rows

            for (int i = 0; i < dataRowCount; i++)
            {
                // Column index 3 corresponds to the fourth column of the main table
                Cell cell = sheet.Cells[startRow + i, startCol + 3];
                Console.WriteLine($"Row {startRow + i + 1} (Cell {cell.Name}): {cell.Formula}");
            }

            // Save the workbook
            string outputPath = "CalculatedColumn_XLookup.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
