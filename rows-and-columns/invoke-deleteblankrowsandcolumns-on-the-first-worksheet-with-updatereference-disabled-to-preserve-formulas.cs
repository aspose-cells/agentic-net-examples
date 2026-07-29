// Title: C# – Delete Blank Rows & Columns on First Worksheet while Preserving Formulas (Aspose.Cells)
// Description: Creates a workbook, adds values and formulas, sets DeleteOptions.UpdateReference = false, then calls DeleteBlankRows and DeleteBlankColumns on the first worksheet to remove empty rows/columns without altering formula references, and saves the file as XLSX.
// Keywords: Aspose.Cells | C# | .NET | DeleteBlankRows | DeleteBlankColumns | DeleteOptions | UpdateReference false | preserve formulas | remove empty rows | remove empty columns | Excel cleanup example
// Common Searches: Aspose.Cells delete blank rows without breaking formulas | DeleteBlankRows UpdateReference false C# | How to keep formula references when removing empty columns in Aspose.Cells | C# example for deleting blank rows and columns in Excel workbook | Aspose.Cells DeleteOptions preserve formulas
// Developer Intent: Remove all empty rows and columns from the first worksheet while ensuring existing formulas keep their original references.
// Use Cases: Clean generated reports by stripping placeholder rows/columns without affecting summary calculations. | Prepare a template workbook for export, eliminating unused cells while keeping dependent formulas intact. | Automate data preprocessing to prune empty rows and columns in large Excel files without breaking cross‑sheet references.
// AI Prompts: Write C# code using Aspose.Cells to delete blank rows and columns with DeleteOptions.UpdateReference set to false. | Explain the effect of DeleteOptions.UpdateReference on formula references when calling DeleteBlankRows and DeleteBlankColumns. | Show how to preserve formulas across multiple worksheets after removing empty rows and columns in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

// Creates a workbook, adds values and formulas, sets DeleteOptions.UpdateReference = false, then calls DeleteBlankRows and DeleteBlankColumns on the first worksheet to remove empty rows/columns without altering formula references, and saves the file as XLSX.
class DeleteBlankRowsAndColumnsExample
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data with blank rows/columns and formulas
        cells["A1"].PutValue(10);
        cells["B1"].PutValue(20);
        cells["C1"].Formula = "=A1+B1";   // Formula referencing A1 and B1

        // Row 2 is intentionally left blank to be removed
        cells["A3"].PutValue(30);
        cells["B3"].PutValue(40);
        cells["C3"].Formula = "=A3+B3";

        // Configure DeleteOptions: do NOT update references (preserve formulas)
        DeleteOptions options = new DeleteOptions
        {
            UpdateReference = false
        };

        // Delete blank rows and columns using the options
        cells.DeleteBlankRows(options);
        cells.DeleteBlankColumns(options);

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
