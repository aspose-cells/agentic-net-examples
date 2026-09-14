// Title: Export all named ranges and their RefersTo formulas to a new audit worksheet with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file, creates a worksheet named 'NamedRangesAudit', and writes each defined name together with its RefersTo formula into two columns using Aspose.Cells. | Generate a method that extracts every defined name and its RefersTo expression from a workbook and writes them to a newly created sheet for verification. | Provide a complete Aspose.Cells example that adds a header row, iterates through the NameCollection, records name and formula pairs, and saves the workbook with the audit sheet.
// Common Searches: Aspose.Cells C# export defined names and formulas to a separate sheet | How to list all named ranges with their RefersTo expressions using Aspose.Cells .NET | Create a verification worksheet for named ranges in an Excel workbook with Aspose.Cells | Save workbook after adding a sheet that contains name and formula columns for each named range
// Tags: export named ranges Aspose.Cells C# | retrieve RefersTo defined names | add overview worksheet Excel | iterate NameCollection workbook | save workbook with additional sheet

using System;
using Aspose.Cells;

// The program loads 'input.xlsx', adds a worksheet called 'NamedRangesAudit', writes header cells, iterates through the workbook's NameCollection, records each defined name and its RefersTo formula, and saves the result as 'output_with_named_ranges_audit.xlsx'.
class ExportNamedRanges
{
    static void Main()
    {
        // Load the source workbook
        Workbook srcWorkbook = new Workbook("input.xlsx");

        // Add a new worksheet for auditing named ranges
        Worksheet auditSheet = srcWorkbook.Worksheets[srcWorkbook.Worksheets.Add()];
        auditSheet.Name = "NamedRangesAudit";

        // Write header row
        auditSheet.Cells[0, 0].PutValue("Name");
        auditSheet.Cells[0, 1].PutValue("Refers To (Formula)");

        // Retrieve the collection of named ranges
        NameCollection names = srcWorkbook.Worksheets.Names;

        // Export each named range and its formula
        int row = 1;
        foreach (Name name in names)
        {
            // Name of the named range
            auditSheet.Cells[row, 0].PutValue(name.Text);

            // Formula that the named range refers to (e.g., =Sheet1!$A$1:$B$10)
            auditSheet.Cells[row, 1].PutValue(name.RefersTo);

            row++;
        }

        // Save the workbook with the audit sheet
        srcWorkbook.Save("output_with_named_ranges_audit.xlsx");
    }
}
