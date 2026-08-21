// Title: Export Named Ranges and Their RefersTo Formulas to an Audit Sheet with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, defines sample named ranges, adds a "NamedRangesAudit" worksheet, writes column headers, iterates the workbook's NameCollection, and records each range's name and RefersTo formula. The columns are auto‑fitted and the file is saved as an Excel audit report.
// Keywords: Aspose.Cells C# | export named ranges | RefersTo formula | list defined names | Excel audit worksheet | retrieve named range addresses | .NET Excel automation | named range reporting | Workbook NameCollection | generate range inventory
// Common Searches: Aspose.Cells list all named ranges .NET | How to export RefersTo formulas to a sheet using C# | Create an audit report of named ranges in Excel with Aspose.Cells | Iterate NameCollection and write names to worksheet | Export defined names to a new worksheet programmatically
// Developer Intent: Create an Excel worksheet that enumerates every named range in a workbook together with its RefersTo expression for auditing or documentation.
// Use Cases: Produce a compliance report that shows all named ranges and their target cells. | Validate that named ranges point to the correct ranges before sharing a workbook. | Provide users with a quick reference sheet of all defined names in a workbook.
// AI Prompts: Generate C# code using Aspose.Cells that adds an audit sheet listing each named range and its RefersTo formula. | Modify the example to also capture the scope (worksheet or workbook) of each named range in the audit report. | Write a reusable method that returns a DataTable with Name and RefersTo columns for all defined names in a Workbook.

using System;
using Aspose.Cells;

namespace NamedRangeAudit
{
    // This example creates a workbook, defines sample named ranges, adds a "NamedRangesAudit" worksheet, writes column headers, iterates the workbook's NameCollection, and records each range's name and RefersTo formula. The columns are auto‑fitted and the file is saved as an Excel audit report.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // OPTIONAL: create some sample named ranges for demo
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Data";
            sheet1.Cells["A1"].PutValue("Item");
            sheet1.Cells["B1"].PutValue("Qty");
            sheet1.Cells["A2"].PutValue("Apple");
            sheet1.Cells["B2"].PutValue(10);
            sheet1.Cells["A3"].PutValue("Banana");
            sheet1.Cells["B3"].PutValue(20);

            // Define two named ranges
            sheet1.Cells.CreateRange("A1:B1").Name = "HeaderRange";
            sheet1.Cells.CreateRange("A2:B3").Name = "DataRange";

            // -------------------------------------------------
            // Create an audit worksheet to list named ranges
            // -------------------------------------------------
            int auditIndex = workbook.Worksheets.Add();
            Worksheet auditSheet = workbook.Worksheets[auditIndex];
            auditSheet.Name = "NamedRangesAudit";

            // Write header row
            auditSheet.Cells["A1"].PutValue("Name");
            auditSheet.Cells["B1"].PutValue("RefersTo");

            // Retrieve all defined names
            NameCollection names = workbook.Worksheets.Names;

            // Iterate through the collection and write details
            int row = 1; // zero‑based index; start after header
            foreach (Name name in names)
            {
                auditSheet.Cells[row, 0].PutValue(name.Text);        // Name text
                auditSheet.Cells[row, 1].PutValue(name.RefersTo);   // Formula (e.g., =Sheet1!$A$1:$B$3)
                row++;
            }

            // Adjust column widths for readability
            auditSheet.AutoFitColumns();

            // Save the workbook to a file
            workbook.Save("NamedRangesAudit.xlsx");
        }
    }
}
