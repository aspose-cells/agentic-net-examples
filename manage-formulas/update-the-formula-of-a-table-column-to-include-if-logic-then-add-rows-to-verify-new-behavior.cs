// Title: Update a table column with an IF formula and programmatically add rows to verify it using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set IF formula for a column and recalculate workbook | How to add new rows with dynamic IF formulas in an Excel sheet using Aspose.Cells | Programmatically update table column formula and evaluate results with Aspose.Cells .NET | C# example for applying conditional calculation to a Total column in Aspose.Cells | Saving an updated workbook with IF logic to XLSX using Aspose.Cells for .NET
// Tags: set IF formula in Aspose.Cells column | add rows with dynamic formulas Aspose.Cells C# | recalculate workbook formulas Aspose.Cells | save workbook as XLSX using Aspose.Cells | update table column formula programmatically Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new workbook, defines a header row, inserts initial data, assigns an IF formula to the Total column that multiplies Quantity by Price only when Quantity is greater than zero, adds two more rows with formulas that reference the newly added rows, forces a full recalculation, and saves the result as UpdatedTableWithIfFormula.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["D1"].PutValue("Total"); // This column will hold the IF formula

            // Add initial data rows
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(5);
            sheet.Cells["C2"].PutValue(2.0);
            sheet.Cells["D2"].Formula = "=IF(B2>0,B2*C2,0)";

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0);
            sheet.Cells["C3"].PutValue(1.5);
            sheet.Cells["D3"].Formula = "=IF(B3>0,B3*C3,0)";

            // Add new rows to verify the new behavior
            // Row 4: Quantity > 0 (should calculate)
            int newRowIndex = sheet.Cells.MaxDataRow + 1;
            sheet.Cells[newRowIndex, 0].PutValue("Orange");   // Product (column A)
            sheet.Cells[newRowIndex, 1].PutValue(3);          // Quantity (column B)
            sheet.Cells[newRowIndex, 2].PutValue(1.8);        // Price (column C)
            sheet.Cells[newRowIndex, 3].Formula = $"=IF(B{newRowIndex + 1}>0,B{newRowIndex + 1}*C{newRowIndex + 1},0)";

            // Row 5: Quantity = 0 (should result in 0)
            newRowIndex = sheet.Cells.MaxDataRow + 1;
            sheet.Cells[newRowIndex, 0].PutValue("Grapes");
            sheet.Cells[newRowIndex, 1].PutValue(0);
            sheet.Cells[newRowIndex, 2].PutValue(2.5);
            sheet.Cells[newRowIndex, 3].Formula = $"=IF(B{newRowIndex + 1}>0,B{newRowIndex + 1}*C{newRowIndex + 1},0)";

            // Recalculate the workbook to ensure formulas are evaluated
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "UpdatedTableWithIfFormula.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
