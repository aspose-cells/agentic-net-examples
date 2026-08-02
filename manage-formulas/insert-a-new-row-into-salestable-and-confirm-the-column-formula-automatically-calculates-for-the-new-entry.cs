// Title: C# – Insert a Row into an Aspose.Cells ListObject (SalesTable) with Automatic Total Formula
// Description: Shows how to create a workbook, define a ListObject named SalesTable (Item, Quantity, Price, Total), assign a column formula (Quantity × Price) for existing rows, programmatically add a new data row, apply the correct formula to the new Total cell, recalculate all formulas, read the computed total, and save the workbook.
// Keywords: Aspose.Cells | C# | .NET | ListObject | Excel table insert row | column formula | auto calculate total | PutCellFormula | PutCellValue | CalculateFormula | programmatic Excel | sales table example | Excel automation
// Common Searches: Aspose.Cells insert row into ListObject | how to keep formulas when adding rows in Aspose.Cells .NET | auto‑calculate column after inserting table row | C# Aspose.Cells ListObject PutCellFormula example | add new product row to Excel table with formula
// Developer Intent: Add a new data row to a ListObject and ensure the Total column formula updates automatically.
// Use Cases: Generate a sales report where each new product entry automatically computes its total price. | Programmatically extend an existing Excel table while preserving financial formulas. | Create dynamic inventory sheets that recalculate derived columns as rows are added via code.
// AI Prompts: Provide C# code to insert a row into an Aspose.Cells ListObject and automatically apply a column formula for the new row. | Show how to add multiple rows to an Excel table with Aspose.Cells without manually setting each formula reference. | Explain the behavior of ListObject.PutCellFormula with relative references after inserting rows in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, define a ListObject named SalesTable (Item, Quantity, Price, Total), assign a column formula (Quantity × Price) for existing rows, programmatically add a new data row, apply the correct formula to the new Total cell, recalculate all formulas, read the computed total, and save the workbook.
    public class InsertRowIntoSalesTable
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ----- Create sample data for the SalesTable -----
                // Header row
                cells["A1"].PutValue("Item");
                cells["B1"].PutValue("Quantity");
                cells["C1"].PutValue("Price");
                cells["D1"].PutValue("Total");

                // First data row
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(10);
                cells["C2"].PutValue(2);
                // Total will be calculated by a formula

                // Second data row
                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(5);
                cells["C3"].PutValue(3);
                // Total will be calculated by a formula

                // ----- Create a ListObject (table) that covers the data range -----
                // The range includes header and the two data rows (A1:D3)
                int tableIndex = sheet.ListObjects.Add("A1", "D3", true);
                ListObject salesTable = sheet.ListObjects[tableIndex];

                // Set the column formula for the "Total" column (index 3) for existing rows
                // Formula: =Quantity*Price  (B column * C column)
                salesTable.PutCellFormula(1, 3, "=B2*C2"); // Row 2 (first data row)
                salesTable.PutCellFormula(2, 3, "=B3*C3"); // Row 3 (second data row)

                // ----- Insert a new row into the table -----
                // The new row will be the third data row (row offset 3 within the table)
                int newRowOffset = 3; // 0‑based offset: 0 = header, 1‑2 = existing data rows

                // Add values for the new row
                salesTable.PutCellValue(newRowOffset, 0, "Cherry");   // Item
                salesTable.PutCellValue(newRowOffset, 1, 8);         // Quantity
                salesTable.PutCellValue(newRowOffset, 2, 4);         // Price

                // Set the formula for the "Total" column of the new row.
                // The worksheet row number for this entry is 4 (A4:D4), so the formula is =B4*C4
                salesTable.PutCellFormula(newRowOffset, 3, "=B4*C4");

                // ----- Calculate formulas and verify the result -----
                workbook.CalculateFormula();

                // Retrieve the calculated total for the newly inserted row
                Cell totalCell = sheet.Cells["D4"]; // D column, fourth row
                Console.WriteLine($"Calculated Total for Cherry: {totalCell.Value}");

                // Save the workbook
                workbook.Save("SalesTableWithInsertedRow.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
