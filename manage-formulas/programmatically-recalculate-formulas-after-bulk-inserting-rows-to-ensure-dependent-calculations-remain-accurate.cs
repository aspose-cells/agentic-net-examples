// Title: Bulk insert rows and recalculate all formulas in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Insert 10 rows starting at row 6 in the first worksheet, then call workbook.CalculateFormula() to refresh every dependent formula. | After adding multiple rows with Worksheet.Cells.InsertRows, trigger a full workbook formula recalculation in C# using Aspose.Cells.
// Common Searches: Aspose.Cells C# insert multiple rows and refresh formulas | recalculate workbook formulas after inserting rows with Aspose.Cells .NET | bulk row insertion affecting dependent calculations in Excel using Aspose.Cells | C# program to add rows and update all formulas in an .xlsx file | how to use CalculateFormula after InsertRows in Aspose.Cells
// Tags: bulk row insertion Aspose.Cells | calculateformula method usage | worksheet cells insertrows example | excel formula recalculation after row addition | c# update dependent formulas Aspose.Cells

using System;
using Aspose.Cells;

// Loads an existing workbook, inserts ten rows at a specified index in the first worksheet, recalculates all formulas to keep dependent calculations accurate, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Define where to insert rows and how many rows to insert
        int insertAtRowIndex = 5;   // zero‑based index (row 6 in Excel)
        int numberOfRows = 10;      // number of rows to insert

        // Bulk insert rows
        worksheet.Cells.InsertRows(insertAtRowIndex, numberOfRows);

        // Recalculate all formulas in the workbook to ensure dependent calculations are updated
        workbook.CalculateFormula();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
