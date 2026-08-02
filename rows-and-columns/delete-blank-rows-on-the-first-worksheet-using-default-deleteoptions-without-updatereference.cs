// Title: C# – Delete blank rows from the first worksheet with Aspose.Cells DeleteBlankRows (default UpdateReference=false)
// Description: Shows how to create a workbook, insert sample data with intentional empty rows, invoke Cells.DeleteBlankRows() using the default DeleteOptions (UpdateReference = false) to purge all blank rows, and save the file as DeletedBlankRows.xlsx. Compatible with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | DeleteBlankRows | blank rows | remove empty rows | UpdateReference false | default DeleteOptions | worksheet rows | Excel automation
// Common Searches: Aspose.Cells delete blank rows C# example | DeleteBlankRows default behavior Aspose.Cells | Remove empty rows from first worksheet without updating references | How to use DeleteBlankRows in Aspose.Cells for .NET | C# code to delete blank rows in Excel workbook
// Developer Intent: Remove every empty row from the first worksheet while leaving existing cell references unchanged.
// Use Cases: Clean imported CSV or data‑feed files by stripping blank rows before analysis. | Prepare a worksheet for PDF or Excel export where empty rows affect pagination. | Normalize report sheets after programmatic data insertion to ensure contiguous rows for charting or pivot tables.
// AI Prompts: Generate C# code that deletes blank rows on the first worksheet using Aspose.Cells with default DeleteOptions and keeps formulas intact. | Explain the effect of UpdateReference = false when calling Cells.DeleteBlankRows in Aspose.Cells and provide a usage snippet. | Create a step‑by‑step tutorial for removing empty rows from a workbook while preserving cell references, using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace DeleteBlankRowsDemo
{
    // Shows how to create a workbook, insert sample data with intentional empty rows, invoke Cells.DeleteBlankRows() using the default DeleteOptions (UpdateReference = false) to purge all blank rows, and save the file as DeletedBlankRows.xlsx. Compatible with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add sample data with intentional blank rows
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Data1");
            // Row 3 is left blank
            cells["A4"].PutValue("Data2");
            // Row 5 is left blank
            cells["A6"].PutValue("Data3");

            // Delete all blank rows using the default DeleteOptions.
            // The default options have UpdateReference = false, which satisfies the requirement.
            cells.DeleteBlankRows();

            // Save the modified workbook
            workbook.Save("DeletedBlankRows.xlsx", SaveFormat.Xlsx);
        }
    }
}
