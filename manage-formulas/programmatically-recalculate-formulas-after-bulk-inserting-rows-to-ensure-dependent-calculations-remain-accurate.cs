// Title: Recalculate all workbook formulas after bulk inserting rows with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# recalculate formulas after inserting multiple rows | bulk row insertion update dependent formulas Aspose.Cells .NET | programmatically refresh Excel formulas after row shift using Aspose.Cells | calculate workbook formulas after InsertRows method in C# | how to keep formulas accurate after adding rows with Aspose.Cells
// Tags: Aspose.Cells InsertRows method formula recalculation | C# recalculate workbook formulas Aspose.Cells | bulk row insertion Excel formulas update | CalculateFormula after row shift Aspose.Cells | Excel workbook formula refresh .NET

using Aspose.Cells;
using System;

// Loads an existing Excel workbook, inserts a block of rows at a specified index using Aspose.Cells, invokes CalculateFormula to refresh all dependent calculations, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can select any worksheet as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Define where to insert rows and how many rows to insert
        int startRow = 5;      // zero‑based index; inserts before row 6 in Excel UI
        int rowsToInsert = 10; // number of rows to insert

        // Bulk insert rows
        sheet.Cells.InsertRows(startRow, rowsToInsert);

        // After inserting rows, recalculate all formulas in the workbook
        // This ensures that any formulas depending on the shifted cells are updated.
        workbook.CalculateFormula();

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
