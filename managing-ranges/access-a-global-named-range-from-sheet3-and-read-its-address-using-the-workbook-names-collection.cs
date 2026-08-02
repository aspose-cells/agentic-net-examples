// Title: Read the address of a global named range on Sheet3 with Aspose.Cells for .NET
// Description: This example creates a workbook, adds a worksheet named "Sheet3", defines a global named range "MyGlobalRange" that points to $A$1:$B$2, retrieves the name from the workbook's Names collection, obtains the underlying Range object, and outputs both the range address and its worksheet before saving the file.
// Keywords: Aspose.Cells | C# | .NET | global named range | workbook Names collection | GetRange | read range address | Sheet3 | example code | retrieve named range
// Common Searches: Aspose.Cells get address of global named range | C# read named range address from workbook | How to access global named range in Aspose.Cells | Retrieve worksheet name of a named range Aspose.Cells | List all named ranges and their addresses using Aspose.Cells
// Developer Intent: The developer wants to obtain the cell address and the worksheet name of a global named range that was defined on Sheet3, using Aspose.Cells for .NET.
// Use Cases: Confirm that a global named range points to the correct cells before data processing. | Log range address and sheet name for debugging or documentation. | Apply formatting, formulas, or data validation dynamically based on a retrieved named range. | Generate migration scripts that need to reference existing named ranges.
// AI Prompts: Generate C# code with Aspose.Cells that enumerates all global named ranges in a workbook and prints each address and its worksheet. | Show how to update the RefersTo property of an existing global named range to a new range on a different worksheet. | Create a reusable method that returns the address string of a specified global named range. | Write a unit test that verifies a global named range "MyGlobalRange" references cells $A$1:$B$2 on Sheet3.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a worksheet named "Sheet3", defines a global named range "MyGlobalRange" that points to $A$1:$B$2, retrieves the name from the workbook's Names collection, obtains the underlying Range object, and outputs both the range address and its worksheet before saving the file.
    public class AccessGlobalNamedRange
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Ensure there is a worksheet named "Sheet3"
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Add a global named range (global = SheetIndex 0)
            int nameIndex = workbook.Worksheets.Names.Add("MyGlobalRange");
            Name globalName = workbook.Worksheets.Names[nameIndex];

            // Define the range on Sheet3 (e.g., A1:B2)
            globalName.RefersTo = $"={sheet3.Name}!$A$1:$B$2";

            // Explicitly set SheetIndex to 0 to make it global (optional, default is global)
            globalName.SheetIndex = 0;

            // Access the global named range via the Names collection using its text
            Name retrievedName = workbook.Worksheets.Names["MyGlobalRange"];
            if (retrievedName != null)
            {
                // Get the Range object that the name refers to
                Aspose.Cells.Range range = retrievedName.GetRange();

                // Read and output the address of the range
                Console.WriteLine("Global named range address: " + range.Address);
                Console.WriteLine("Defined on worksheet: " + range.Worksheet.Name);
            }
            else
            {
                Console.WriteLine("Named range 'MyGlobalRange' not found.");
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("AccessGlobalNamedRange.xlsx");
        }
    }
}
