using Aspose.Cells;
using System;

class ExportNamedRanges
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Sample data and named ranges (for demonstration)
        // -------------------------------------------------
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Populate some cells
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["B1"].PutValue(30);
        sheet.Cells["B2"].PutValue(40);

        // Create named ranges
        sheet.Cells.CreateRange("A1:B2").Name = "MyRange";
        sheet.Cells.CreateRange("A1").Name = "SingleCell";

        // -------------------------------------------------
        // Create an audit worksheet to list named ranges
        // -------------------------------------------------
        int auditSheetIndex = workbook.Worksheets.Add();
        Worksheet auditSheet = workbook.Worksheets[auditSheetIndex];
        auditSheet.Name = "NamedRangesAudit";

        // Write header
        auditSheet.Cells["A1"].PutValue("Name");
        auditSheet.Cells["B1"].PutValue("RefersTo");

        // Iterate through all defined names and export their details
        int currentRow = 1; // zero‑based index; row 1 is the second row
        foreach (Name definedName in workbook.Worksheets.Names)
        {
            // Column A: Name text
            auditSheet.Cells[currentRow, 0].PutValue(definedName.Text);
            // Column B: Formula the name refers to (e.g., "=Sheet1!$A$1:$B$2")
            auditSheet.Cells[currentRow, 1].PutValue(definedName.RefersTo);
            currentRow++;
        }

        // -------------------------------------------------
        // Save the workbook with the audit sheet
        // -------------------------------------------------
        workbook.Save("NamedRangesAudit.xlsx");
    }
}