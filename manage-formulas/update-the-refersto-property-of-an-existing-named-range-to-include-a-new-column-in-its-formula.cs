// Title: C# – Expand a Named Range RefersTo to Include an Additional Column with Aspose.Cells
// Description: Creates a workbook, defines a named range that points to column A, then updates the Name.RefersTo property so the range covers columns A‑B, prints the new address, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | named range | RefersTo | update range address | add column to named range | Excel automation | workbook manipulation
// Common Searches: Aspose.Cells change RefersTo address | C# update named range to include another column | How to modify a named range in Aspose.Cells .NET | Expand named range columns Aspose.Cells | Add column to existing named range programmatically
// Developer Intent: Modify the RefersTo string of an existing named range so it spans an extra column.
// Use Cases: Adjust a named range after inserting a new data column so dependent formulas automatically include the new data. | Synchronize a named range with a shifting table layout before exporting the workbook for downstream processing. | Programmatically broaden a named range definition to cover additional columns when generating reports.
// AI Prompts: Generate C# code using Aspose.Cells to extend a named range RefersTo from columns A‑B to C‑D. | Explain how to retrieve, validate, and rewrite the RefersTo string of a named range in Aspose.Cells for .NET. | Show a method that expands a named range's RefersTo property based on the last used column in a worksheet.

using System;
using Aspose.Cells;

// Creates a workbook, defines a named range that points to column A, then updates the Name.RefersTo property so the range covers columns A‑B, prints the new address, and saves the file.
class UpdateNamedRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Populate sample data in columns A and B (rows 1-3)
        sheet.Cells["A1"].PutValue(1);
        sheet.Cells["A2"].PutValue(2);
        sheet.Cells["A3"].PutValue(3);
        sheet.Cells["B1"].PutValue(10);
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["B3"].PutValue(30);

        // Add a named range that initially refers only to column A
        int nameIndex = workbook.Worksheets.Names.Add("MyRange");
        Name myRange = workbook.Worksheets.Names[nameIndex];
        myRange.RefersTo = "=Sheet1!$A$1:$A$3";

        // Update the RefersTo property to include the new column B
        myRange.RefersTo = "=Sheet1!$A$1:$B$3";

        // Output the updated RefersTo formula for verification
        Console.WriteLine("Updated RefersTo: " + myRange.RefersTo);

        // Save the workbook
        workbook.Save("UpdatedNamedRange.xlsx");
    }
}
