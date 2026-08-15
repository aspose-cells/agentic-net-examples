// Title: C# – Set an absolute RefersTo address for a named range in Aspose.Cells to prevent shifting
// Description: Creates a workbook, adds data to Sheet1, defines a named range "MyAbsoluteRange" and assigns its RefersTo property with an absolute A1‑style address ("=Sheet1!$A$1:$A$3") using SetRefersTo, ensuring the range stays fixed when rows or columns are inserted, then saves the file.
// Keywords: Aspose.Cells | named range absolute address | SetRefersTo | C# | RefersTo property | prevent range shift | A1 style reference
// Common Searches: Aspose.Cells set RefersTo absolute address C# | prevent named range from moving when inserting rows Aspose.Cells | SetRefersTo parameters isR1C1 false isLocal false | how to create fixed named range Aspose.Cells | absolute reference in named range Aspose.Cells .NET
// Developer Intent: Define a named range with an absolute A1‑style address so it remains unchanged after inserting rows or columns.
// Use Cases: Maintain a constant reference to cells A1:A3 for formulas, charts, or data validation. | Export workbooks where downstream processes rely on stable named ranges. | Create templates that preserve key ranges despite user edits or automated row insertions.
// AI Prompts: Write C# code with Aspose.Cells that creates a named range using an absolute RefersTo address that does not shift when rows are added. | Explain the purpose of the isR1C1 and isLocal parameters in SetRefersTo when assigning an absolute address to a named range.

using System;
using Aspose.Cells;

namespace AsposeCellsAbsoluteNamedRange
{
    // Creates a workbook, adds data to Sheet1, defines a named range "MyAbsoluteRange" and assigns its RefersTo property with an absolute A1‑style address ("=Sheet1!$A$1:$A$3") using SetRefersTo, ensuring the range stays fixed when rows or columns are inserted, then saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate some sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Add a named range to the workbook
            int nameIndex = workbook.Worksheets.Names.Add("MyAbsoluteRange");
            Name namedRange = workbook.Worksheets.Names[nameIndex];

            // Set the RefersTo property using an absolute address.
            // The address is absolute ($ signs) so it will not shift when rows/columns are inserted.
            // Using SetRefersTo with isR1C1 = false (A1 style) and isLocal = false (invariant locale).
            namedRange.SetRefersTo("=Sheet1!$A$1:$A$3", false, false);

            // Optionally, you could also assign directly:
            // namedRange.RefersTo = "=Sheet1!$A$1:$A$3";

            // Save the workbook to a file
            workbook.Save("AbsoluteNamedRange.xlsx");
        }
    }
}
