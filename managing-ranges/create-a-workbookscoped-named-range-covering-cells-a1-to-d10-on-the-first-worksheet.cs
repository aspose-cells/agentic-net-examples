// Title: Aspose.Cells .NET – Create a workbook‑scoped named range A1:D10 on the first sheet
// Description: C# example that creates a new Workbook, adds a workbook‑scoped named range called "MyRange" covering cells A1:D10 on the first worksheet, sets the RefersTo reference, and saves the file as WorkbookScopedNamedRange.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells named range .NET | workbook scoped range C# | add named range A1:D10 | RefersTo property Aspose.Cells | save workbook with named range | Aspose.Cells example C# | Excel named range programmatically
// Common Searches: Aspose.Cells add workbook scoped named range | C# create named range A1:D10 Aspose.Cells | set RefersTo for named range Aspose.Cells .NET | how to define named range on first worksheet Aspose | Aspose.Cells save workbook with named range
// Developer Intent: Add a workbook‑scoped named range named "MyRange" that points to A1:D10 on the first worksheet and persist the workbook.
// Use Cases: Reference a fixed data block in formulas across multiple sheets. | Apply data validation or conditional formatting using a reusable range. | Enable external tools to locate a specific table by name when the workbook is shared.
// AI Prompts: Generate C# code with Aspose.Cells to create a workbook‑scoped named range covering A1:D10 on the first sheet and save the workbook. | Show how to modify the RefersTo property to move the named range to another worksheet or a different cell range. | Provide C# error‑handling for adding a named range when the name already exists in the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeDemo
{
    // C# example that creates a new Workbook, adds a workbook‑scoped named range called "MyRange" covering cells A1:D10 on the first worksheet, sets the RefersTo reference, and saves the file as WorkbookScopedNamedRange.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the address of the range A1:D10
            string rangeAddress = "$A$1:$D$10";

            // Add a workbook‑scoped named range called "MyRange"
            // Index of the newly added name
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            // Retrieve the Name object
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            // Set the reference to the range on the first worksheet
            // The reference must start with '=' and include the sheet name
            namedRange.RefersTo = $"={sheet.Name}!{rangeAddress}";

            // (Optional) Verify that the name was created
            Console.WriteLine($"Created named range '{namedRange.Text}' referring to {namedRange.RefersTo}");

            // Save the workbook to a file
            workbook.Save("WorkbookScopedNamedRange.xlsx");
        }
    }
}
