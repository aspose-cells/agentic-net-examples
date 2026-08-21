// Title: Find cells with formulas that reference a named range using Aspose.Cells FindOptions (C#)
// Description: This example creates a workbook, defines a named range "MyRange" (A1:A3), adds formulas that use the range, and demonstrates how to configure FindOptions (LookInType.OnlyFormulas, LookAtType.Contains) to locate every cell whose formula contains the named range. The code iterates through the matches, prints each cell address and formula, and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | FindOptions | OnlyFormulas | LookAtType.Contains | named range | formula search | cell find | workbook audit | Aspose.Cells Find method
// Common Searches: Aspose.Cells find cells referencing a named range | FindOptions OnlyFormulas example C# | search formulas for a specific name in Aspose.Cells | how to locate cells that use a named range with Aspose.Cells | C# code to find formulas containing MyRange
// Developer Intent: Identify all cells whose formulas reference a particular named range.
// Use Cases: Audit a workbook to list every cell that uses a given named range before restructuring data. | Validate consistency of named‑range references across multiple worksheets. | Programmatically replace an outdated named range with a new one throughout a workbook.
// AI Prompts: Generate C# code with Aspose.Cells that finds all cells whose formulas contain the named range "MyRange" and replaces it with "NewRange". | Explain how to set FindOptions to search only within formulas for a specific string in Aspose.Cells. | Create a method that returns a list of cell addresses that reference a supplied named range using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace FindFormulasReferencingNamedRange
{
    // This example creates a workbook, defines a named range "MyRange" (A1:A3), adds formulas that use the range, and demonstrates how to configure FindOptions (LookInType.OnlyFormulas, LookAtType.Contains) to locate every cell whose formula contains the named range. The code iterates through the matches, prints each cell address and formula, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data in A1:A3
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Create a named range called "MyRange" that refers to A1:A3
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name myRange = workbook.Worksheets.Names[nameIndex];
            myRange.RefersTo = "=Sheet1!$A$1:$A$3";

            // Add formulas that reference the named range
            cells["B1"].Formula = "=SUM(MyRange)";
            cells["B2"].Formula = "=AVERAGE(MyRange)";
            cells["C1"].Formula = "=MAX(MyRange)";
            // A formula that does NOT reference the named range (for contrast)
            cells["D1"].Formula = "=SUM(A1:A3)";

            // Set up FindOptions to search only within formulas and look for the name "MyRange"
            FindOptions options = new FindOptions
            {
                LookInType = LookInType.OnlyFormulas,   // Search only formula text
                LookAtType = LookAtType.Contains        // Name can appear anywhere in the formula
            };

            // Perform the first search
            Cell previous = null;
            Cell found = sheet.Cells.Find("MyRange", previous, options);

            Console.WriteLine("Cells whose formulas reference the named range \"MyRange\":");
            while (found != null)
            {
                Console.WriteLine($"- {found.Name} : {found.Formula}");
                // Continue searching from the cell after the current one
                previous = found;
                found = sheet.Cells.Find("MyRange", previous, options);
            }

            // Save the workbook (optional, demonstrates lifecycle rule usage)
            workbook.Save("FormulasReferencingMyRange.xlsx");
        }
    }
}
