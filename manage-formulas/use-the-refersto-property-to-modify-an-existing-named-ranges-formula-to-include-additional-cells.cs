// Title: Extend a Named Range with Additional Cells Using RefersTo in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to read a named range's RefersTo string, append a new address (e.g., Data!$B$1:$B$3), assign the combined expression back to the Name object, and use the expanded multi‑area range in a SUM formula. The workbook is calculated and saved as ModifiedNamedRange.xlsx.
// Keywords: Aspose.Cells C# named range | RefersTo property extend range | multi‑area named range Aspose.Cells | add cells to existing named range .NET | update named range formula Aspose | dynamic range modification Aspose.Cells
// Common Searches: how to add a second area to a named range in Aspose.Cells | extend named range RefersTo C# Aspose.Cells | combine multiple ranges into one named range Aspose | modify RefersTo string programmatically Aspose.Cells | multi‑area named range example Aspose.Cells
// Developer Intent: Modify an existing named range so it references additional cells and apply the updated range in dependent formulas.
// Use Cases: Create a multi‑area named range for consolidated calculations like SUM or AVERAGE across non‑contiguous columns. | Adjust a data source range dynamically before generating charts or pivot tables. | Programmatically expand a named range based on user input or runtime data without recreating the name.
// AI Prompts: Provide C# code that uses Aspose.Cells to append a new address to an existing named range via the RefersTo property and recalculate dependent formulas. | Write a reusable method that takes a Workbook, a named range name, and an extra address, then updates the RefersTo string to include the new area while preserving existing references.

using System;
using Aspose.Cells;

// Demonstrates how to read a named range's RefersTo string, append a new address (e.g., Data!$B$1:$B$3), assign the combined expression back to the Name object, and use the expanded multi‑area range in a SUM formula. The workbook is calculated and saved as ModifiedNamedRange.xlsx.
class ModifyNamedRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Populate some data in column A
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);

        // Create a named range that initially refers to A1:A3
        int nameIdx = workbook.Worksheets.Names.Add("MyRange");
        Name myRange = workbook.Worksheets.Names[nameIdx];
        myRange.RefersTo = "=Data!$A$1:$A$3";

        // Extend the named range to also include B1:B3
        // Build a new RefersTo string that combines the existing area with the new one
        string currentRef = myRange.RefersTo.TrimStart('='); // "Data!$A$1:$A$3"
        string extendedRef = $"={currentRef},Data!$B$1:$B$3";
        myRange.RefersTo = extendedRef;

        // Use the extended named range in a formula
        sheet.Cells["C1"].Formula = "=SUM(MyRange)";
        workbook.CalculateFormula();

        // Display the updated RefersTo and the calculated sum
        Console.WriteLine("Extended RefersTo: " + myRange.RefersTo);
        Console.WriteLine("Sum of extended range: " + sheet.Cells["C1"].Value);

        // Save the workbook
        workbook.Save("ModifiedNamedRange.xlsx");
    }
}
