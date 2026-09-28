// Title: Insert a row in an Aspose.Cells worksheet and automatically recalculate dependent formulas using Workbook.CalculateFormula (C#)
// AI Prompts: Insert a new row at a specific index in an Aspose.Cells worksheet, add data to the inserted cells, and invoke Workbook.CalculateFormula to refresh all formulas. | Show how to update a SUM formula after shifting rows by inserting a row and calling the calculation engine in C#. | Provide complete C# code that inserts a row, modifies cell values, recalculates formulas, and saves the workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# insert row and recalculate formulas example | How to refresh SUM formula after adding a row with Aspose.Cells | Workbook.CalculateFormula after row insertion in .NET | Update dependent formulas automatically when inserting rows using Aspose.Cells | C# code to insert row and trigger formula recalculation in Excel file
// Tags: add new row Aspose.Cells C# | Workbook.CalculateFormula recalculate | refresh dependent formulas Aspose.Cells | update SUM after row insertion | save workbook after calculation Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a workbook, fills cells A1‑A3, sets B1 to =SUM(A1:A3), inserts a row at index 1, adds a value to the new A2, calls Workbook.CalculateFormula to update the SUM, prints the refreshed B1 value, and saves the file as InsertedRow.xlsx.
class InsertRowAndRecalculate
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Initial data
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["A3"].PutValue(30);
        // Formula that sums the three values
        cells["B1"].Formula = "=SUM(A1:A3)";

        // Insert a new row at index 1 (between A1 and A2)
        cells.InsertRow(1);
        // Add a value to the newly inserted row
        cells["A2"].PutValue(15); // original A2 and A3 shift down

        // Recalculate formulas after the insertion
        workbook.CalculateFormula();

        // Display the updated result of the formula
        Console.WriteLine("B1 after recalculation: " + cells["B1"].Value);

        // Save the workbook (optional)
        workbook.Save("InsertedRow.xlsx");
    }
}
