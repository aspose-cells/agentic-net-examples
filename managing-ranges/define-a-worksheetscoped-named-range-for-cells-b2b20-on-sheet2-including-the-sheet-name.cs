// Title: Define a worksheet‑scoped named range B2:B20 on Sheet2 with Aspose.Cells for .NET
// Description: This example shows how to create a new workbook, add a worksheet named "Sheet2", and define a worksheet‑scoped named range called "MyRange" that points to cells B2:B20 on that sheet. The range is scoped by setting the RefersTo address with the sheet name and assigning the SheetIndex, then the workbook is saved as WorksheetScopedNamedRange.xlsx.
// Keywords: Aspose.Cells | .NET | C# | worksheet scoped named range | named range with sheet name | RefersTo property | SheetIndex | Excel named range programmatically | B2:B20 range | Aspose.Cells Names collection
// Common Searches: Aspose.Cells create worksheet scoped named range C# | set named range RefersTo with sheet name Aspose.Cells | define named range B2:B20 on specific worksheet .NET | how to limit named range scope to a single sheet using Aspose.Cells | C# code for worksheet‑level named range in Excel
// Developer Intent: Create a worksheet‑scoped named range for cells B2:B20 on Sheet2.
// Use Cases: Use the range in data‑validation lists that should only apply to Sheet2. | Reference the range in formulas on Sheet2 without hard‑coding cell addresses. | Link a chart series on Sheet2 to a named range that is isolated from other sheets.
// AI Prompts: Generate C# code with Aspose.Cells that adds a worksheet‑scoped named range "MyRange" for B2:B20 on a worksheet named "Sheet2" and saves the file. | Explain how to change the scope of an existing named range to a specific worksheet using Aspose.Cells for .NET. | Show how to read the RefersTo address of a worksheet‑scoped named range and use it in a formula programmatically.

using System;
using Aspose.Cells;

// This example shows how to create a new workbook, add a worksheet named "Sheet2", and define a worksheet‑scoped named range called "MyRange" that points to cells B2:B20 on that sheet. The range is scoped by setting the RefersTo address with the sheet name and assigning the SheetIndex, then the workbook is saved as WorksheetScopedNamedRange.xlsx.
class DefineWorksheetScopedNamedRange
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a worksheet named "Sheet2"
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        // Add a worksheet‑scoped named range for B2:B20 on Sheet2
        // 1. Add the name to the workbook's Names collection
        int nameIndex = workbook.Worksheets.Names.Add("MyRange");
        Name namedRange = workbook.Worksheets.Names[nameIndex];

        // 2. Set the reference to the desired cells, including the sheet name
        namedRange.RefersTo = $"={sheet2.Name}!$B$2:$B$20";

        // 3. Set the scope to the worksheet (one‑based sheet index)
        namedRange.SheetIndex = workbook.Worksheets.IndexOf(sheet2) + 1;

        // Save the workbook (adjust the path as needed)
        workbook.Save("WorksheetScopedNamedRange.xlsx");
    }
}
