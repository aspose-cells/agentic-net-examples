using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class ListObjectCalculatedColumnWithXLookup
    {
        public static void Run()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Populate main table data ----------
                // Headers
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["C1"].PutValue("Category"); // Placeholder for calculated column

                // Data rows
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(12);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(15);

                // ---------- Populate lookup table data ----------
                // Headers
                sheet.Cells["E1"].PutValue("Item");
                sheet.Cells["F1"].PutValue("Category");

                // Data rows
                sheet.Cells["E2"].PutValue("Apple");
                sheet.Cells["F2"].PutValue("Fruit");
                sheet.Cells["E3"].PutValue("Banana");
                sheet.Cells["F3"].PutValue("Fruit");
                sheet.Cells["E4"].PutValue("Carrot");
                sheet.Cells["F4"].PutValue("Vegetable");

                // ---------- Create ListObjects (tables) ----------
                // Main table (A1:C4)
                int mainTableIndex = sheet.ListObjects.Add("A1", "C4", true);
                ListObject mainTable = sheet.ListObjects[mainTableIndex];
                // Set table name (DisplayName is the correct property)
                mainTable.DisplayName = "MainTable";

                // Lookup table (E1:F4)
                int lookupTableIndex = sheet.ListObjects.Add("E1", "F4", true);
                ListObject lookupTable = sheet.ListObjects[lookupTableIndex];
                lookupTable.DisplayName = "LookupTable";

                // ---------- Add calculated column using XLOOKUP ----------
                // The calculated column is the third column (index 2) of the main table
                ListColumn categoryColumn = mainTable.ListColumns[2];
                categoryColumn.Name = "Category";

                // Structured reference [@Item] refers to the Item value of the current row.
                // Lookup range uses absolute A1 notation.
                string xlookupFormula = "=XLOOKUP([@Item],LookupTable[Item],LookupTable[Category],\"NotFound\")";

                // Set custom calculated formula for the entire column.
                categoryColumn.SetCustomCalculatedFormula(xlookupFormula, false, false);

                // ---------- Calculate formulas ----------
                workbook.CalculateFormula();

                // ---------- Verify automatic propagation ----------
                Console.WriteLine("Calculated Category values:");
                for (int row = 1; row <= 3; row++) // Data rows start at index 1 (row 2 in worksheet)
                {
                    // Cells are zero‑based; column C is index 2.
                    Cell cell = sheet.Cells[row, 2];
                    Console.WriteLine($"Row {row + 1}: {cell.StringValue}");
                }

                // ---------- Save the workbook ----------
                string outputPath = "ListObjectWithXLookupCalculatedColumn.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ListObjectCalculatedColumnWithXLookup.Run();
        }
    }
}