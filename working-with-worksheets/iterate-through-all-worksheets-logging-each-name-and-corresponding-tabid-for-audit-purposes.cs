// Title: Audit Worksheet Names and TabIds with Aspose.Cells for .NET
// Description: Shows how to loop through all worksheets in an Aspose.Cells workbook, output each worksheet's Name and TabId, and optionally save the file—ideal for creating audit logs, validating identifiers, or generating reports.
// Keywords: Aspose.Cells | .NET | C# | Worksheet TabId | enumerate worksheets | audit worksheet identifiers | log worksheet names | retrieve TabId | worksheet metadata | Aspose.Cells API
// Common Searches: Aspose.Cells get TabId for each worksheet C# | iterate worksheets and log names Aspose.Cells | how to audit worksheet identifiers with Aspose.Cells | retrieve worksheet TabId property .NET | list all worksheet names and TabIds Aspose.Cells
// Developer Intent: The developer needs to enumerate every worksheet in a workbook and record its Name and TabId for auditing, validation, or reporting purposes.
// Use Cases: Create an audit trail of worksheet identifiers before sharing a workbook. | Verify that TabId values are unique to avoid navigation conflicts. | Generate a version‑control report showing worksheet order and IDs.
// AI Prompts: Write a reusable method that returns a dictionary of worksheet names and TabId values using Aspose.Cells for .NET. | Add robust error handling so that a null workbook or missing worksheets are logged to a file instead of the console. | Show how to filter worksheets by a naming pattern while still logging their TabId for selective auditing.

using System;
using Aspose.Cells;

// Shows how to loop through all worksheets in an Aspose.Cells workbook, output each worksheet's Name and TabId, and optionally save the file—ideal for creating audit logs, validating identifiers, or generating reports.
class Program
{
    static void Main()
    {
        // Create a new workbook (you can replace this with loading an existing file if needed)
        Workbook workbook = new Workbook();

        // Example: add/rename worksheets to demonstrate the audit logging
        workbook.Worksheets[0].Name = "SheetOne";
        workbook.Worksheets.Add("SheetTwo");
        workbook.Worksheets.Add("SheetThree");

        // Iterate through all worksheets and log their Name and TabId
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet Name: {sheet.Name}, TabId: {sheet.TabId}");
        }

        // Save the workbook (optional, based on your workflow)
        workbook.Save("AuditLogWorkbook.xlsx");
    }
}
