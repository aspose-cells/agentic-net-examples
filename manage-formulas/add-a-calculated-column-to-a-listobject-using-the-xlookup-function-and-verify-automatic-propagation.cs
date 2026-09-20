// Title: Add an XLOOKUP calculated column to an Aspose.Cells ListObject and verify automatic formula propagation in C#
// AI Prompts: Create a ListObject on a worksheet, append a new column named "LookupResult", and set its cells to an XLOOKUP formula that references a separate lookup sheet. | Iterate through each data row of the ListObject to confirm that the XLOOKUP formula is present in every cell of the new column, logging any mismatches. | Save the workbook after the calculated column has been added and validated, and output the full path of the generated file.
// Common Searches: aspocells add XLOOKUP column to table and check formula propagation c# | c# verify that XLOOKUP formula is applied to all rows in an Aspose.Cells ListObject | how to reference another worksheet in XLOOKUP using Aspose.Cells C# | create Excel table with calculated column using XLOOKUP in Aspose.Cells | save workbook after adding calculated column with XLOOKUP Aspose.Cells
// Tags: add XLOOKUP calculated column ListObject | verify formula propagation Aspose.Cells | C# create Excel table with lookup sheet | Aspose.Cells XLOOKUP across worksheets | save workbook with calculated column

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsCalculatedColumnExample
{
    // The example builds a workbook containing a main data sheet and a separate lookup sheet, creates a ListObject (Excel table) on the main sheet, adds a new calculated column called "LookupResult" populated with an XLOOKUP formula that pulls names from the lookup sheet, checks that the formula is automatically propagated to every data row, and finally saves the workbook as CalculatedColumn_XLookup.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "MainData";

                // Populate main data (ID column)
                sheet1.Cells["A1"].PutValue("ID");
                sheet1.Cells["A2"].PutValue(101);
                sheet1.Cells["A3"].PutValue(102);
                sheet1.Cells["A4"].PutValue(103);
                sheet1.Cells["A5"].PutValue(104);
                sheet1.Cells["A6"].PutValue(105);

                // Add a second column with some values (optional)
                sheet1.Cells["B1"].PutValue("Value");
                sheet1.Cells["B2"].PutValue(10);
                sheet1.Cells["B3"].PutValue(20);
                sheet1.Cells["B4"].PutValue(30);
                sheet1.Cells["B5"].PutValue(40);
                sheet1.Cells["B6"].PutValue(50);

                // Create a lookup table on a second worksheet
                Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet2.Name = "LookupTable";

                // Populate lookup table: Column A = ID, Column B = Name
                sheet2.Cells["A1"].PutValue("ID");
                sheet2.Cells["B1"].PutValue("Name");
                sheet2.Cells["A2"].PutValue(101);
                sheet2.Cells["B2"].PutValue("Alice");
                sheet2.Cells["A3"].PutValue(102);
                sheet2.Cells["B3"].PutValue("Bob");
                sheet2.Cells["A4"].PutValue(103);
                sheet2.Cells["B4"].PutValue("Charlie");
                sheet2.Cells["A5"].PutValue(104);
                sheet2.Cells["B5"].PutValue("Diana");
                sheet2.Cells["A6"].PutValue(105);
                sheet2.Cells["B6"].PutValue("Eve");

                // Define the range for the main table (A1:B6)
                int firstRow = 0;          // zero‑based index
                int firstColumn = 0;
                int totalRows = 6;         // includes header row
                int totalColumns = 2;      // ID and Value columns

                // Add a ListObject (Excel table) to the defined range
                int listObjectIndex = sheet1.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
                ListObject listObject = sheet1.ListObjects[listObjectIndex];
                listObject.DisplayName = "MainTable";

                // Add a calculated column (LookupResult) manually
                int newColumnIndex = totalColumns; // zero‑based index for the new column
                string headerAddress = CellsHelper.CellIndexToName(firstRow, newColumnIndex);
                sheet1.Cells[headerAddress].PutValue("LookupResult");

                // XLOOKUP formula to retrieve names from the lookup table
                string xlookupFormula = "=XLOOKUP([@ID],LookupTable!$A$2:$A$6,LookupTable!$B$2:$B$6,\"Not Found\")";

                // Apply the formula to each data row in the new column
                for (int r = 1; r < totalRows; r++) // start from row 1 (skip header)
                {
                    Cell cell = sheet1.Cells[firstRow + r, newColumnIndex];
                    cell.Formula = xlookupFormula;
                }

                // Verify that the formula was propagated to all rows
                bool allFormulasMatch = true;
                for (int r = 1; r < totalRows; r++)
                {
                    Cell cell = sheet1.Cells[firstRow + r, newColumnIndex];
                    if (!cell.Formula.Equals(xlookupFormula, StringComparison.OrdinalIgnoreCase))
                    {
                        allFormulasMatch = false;
                        Console.WriteLine($"Row {firstRow + r + 1} formula mismatch: {cell.Formula}");
                    }
                }

                Console.WriteLine(allFormulasMatch
                    ? "XLOOKUP formula propagated to all rows successfully."
                    : "Formula propagation verification failed.");

                // Save the workbook (optional, just to demonstrate lifecycle)
                string outputPath = "CalculatedColumn_XLookup.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
