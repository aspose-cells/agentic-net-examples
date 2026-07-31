// Title: C# – Rename Active Worksheet to QuarterlyReport with Aspose.Cells (Data Preserved)
// Description: Demonstrates how to access the active worksheet in a new Workbook, change its Name to "QuarterlyReport", and save the file while keeping all existing cell values intact.
// Keywords: Aspose.Cells rename worksheet C# | active sheet rename .NET | preserve cell values Aspose | set worksheet name programmatically | Aspose.Cells workbook rename | C# worksheet name change
// Common Searches: Aspose.Cells rename active sheet | C# change worksheet name without losing data | how to set worksheet name in Aspose.Cells | rename worksheet to QuarterlyReport Aspose | preserve cell data when renaming sheet Aspose.Cells
// Developer Intent: Rename the currently active worksheet to "QuarterlyReport" while keeping all existing cell contents.
// Use Cases: Rename the default sheet of a freshly created workbook before exporting to a client‑specific filename. | Update the sheet name after populating data to reflect the reporting period without altering any cell values. | Automate sheet renaming in a batch process, ensuring data integrity across multiple workbooks.
// AI Prompts: Provide C# code using Aspose.Cells to rename the active worksheet to a given name while preserving all cell data. | Explain the steps to rename a worksheet in Aspose.Cells and verify that cell values are retained after saving the workbook. | Create a reusable method that takes a Workbook object and a new sheet name, renames the active worksheet, and returns the updated workbook.

using System;
using Aspose.Cells;

// Demonstrates how to access the active worksheet in a new Workbook, change its Name to "QuarterlyReport", and save the file while keeping all existing cell values intact.
class RenameActiveWorksheet
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();

        // Access the currently active worksheet
        Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // Example data to show that cell contents are preserved after renaming
        activeSheet.Cells["A1"].PutValue("Sample Data");

        // Rename the active worksheet to "QuarterlyReport"
        activeSheet.Name = "QuarterlyReport";

        // Save the workbook (saving rule)
        workbook.Save("QuarterlyReport.xlsx", SaveFormat.Xlsx);
    }
}
