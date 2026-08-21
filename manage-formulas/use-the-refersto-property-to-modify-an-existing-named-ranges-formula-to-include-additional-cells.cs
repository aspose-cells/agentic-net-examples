// Title: C# – Extend a Named Range Using RefersTo in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, define a named range, update its RefersTo property to include more cells, apply the range in a SUM formula, recalculate, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | named range | RefersTo | modify range | extend range | sample code | GitHub example | worksheet formula
// Common Searches: Aspose.Cells change RefersTo C# | extend named range programmatically Aspose.Cells | update named range formula .NET | add cells to existing named range Aspose | RefersTo property example GitHub
// Developer Intent: Show how to programmatically modify the RefersTo property of an existing named range to cover additional cells.
// Use Cases: Increase the scope of a named range before using it in aggregate functions like SUM or AVERAGE. | Adjust a range dynamically as new data rows are appended to a worksheet. | Reuse an expanded named range across multiple formulas without recreating it.
// AI Prompts: Write C# code that reads a named range's current RefersTo, appends extra cell addresses, and updates the property using Aspose.Cells. | Explain step‑by‑step how to extend a named range's RefersTo to a variable size and recalculate dependent formulas in Aspose.Cells for .NET. | Provide a complete example that modifies a named range, then uses it in different worksheet formulas such as SUM, AVERAGE, and COUNT.

using System;
using Aspose.Cells;

namespace AsposeCellsRefersToDemo
{
    // Demonstrates how to create a workbook, define a named range, update its RefersTo property to include more cells, apply the range in a SUM formula, recalculate, and save the file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate some sample data in cells A1:A5
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1); // Column 0 = "A"
            }

            // Add a named range that initially refers to A1:A3
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name myRange = workbook.Worksheets.Names[nameIndex];
            myRange.RefersTo = "=Sheet1!$A$1:$A$3";

            Console.WriteLine("Original RefersTo: " + myRange.RefersTo);

            // Modify the RefersTo property to include additional cells (A4:A5)
            // The new formula should reference A1:A5
            myRange.RefersTo = "=Sheet1!$A$1:$A$5";

            Console.WriteLine("Updated RefersTo: " + myRange.RefersTo);

            // Use the named range in a formula to verify it works
            sheet.Cells["B1"].Formula = "=SUM(MyRange)";
            workbook.CalculateFormula();

            Console.WriteLine("SUM of MyRange (A1:A5): " + sheet.Cells["B1"].Value);

            // Save the workbook
            workbook.Save("RefersToModified.xlsx");
        }
    }
}
