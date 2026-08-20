// Title: C# – List Workbook‑Scoped Named Ranges with Aspose.Cells and Print to Console
// Description: Shows how to add global (workbook‑scoped) named ranges, filter them with NameScopeType.Workbook, and write each name to the console using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | workbook scoped named ranges | NameScopeType.Workbook | list named ranges | filter global names | enumerate named ranges | console output | Excel automation
// Common Searches: Aspose.Cells list workbook scoped names C# | filter named ranges by scope Aspose.Cells | get global named ranges .NET | enumerate named ranges console Aspose | NameScopeType.Workbook example
// Developer Intent: Retrieve all workbook‑level named ranges from a spreadsheet and display their identifiers.
// Use Cases: Verify that required global named ranges exist before running calculations. | Generate a documentation report that lists every workbook‑scoped named range. | Log workbook‑scoped names for debugging when loading or modifying a workbook.
// AI Prompts: Write C# code with Aspose.Cells that lists all workbook‑scoped named ranges and saves the names to a text file. | Show how to differentiate worksheet‑scoped and workbook‑scoped named ranges and retrieve each group separately. | Provide an example that adds a new workbook‑scoped named range, then enumerates and prints all workbook‑scoped names.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to add global (workbook‑scoped) named ranges, filter them with NameScopeType.Workbook, and write each name to the console using Aspose.Cells for .NET.
    public class ListWorkbookScopedNames
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Add some workbook‑scoped named ranges for demonstration
                int idx1 = workbook.Worksheets.Names.Add("GlobalRange1");
                workbook.Worksheets.Names[idx1].RefersTo = "=Sheet1!$A$1:$A$5";

                int idx2 = workbook.Worksheets.Names.Add("GlobalRange2");
                workbook.Worksheets.Names[idx2].RefersTo = "=Sheet1!$B$1:$B$5";

                // Retrieve only the workbook‑scoped names using the Filter method
                Name[] workbookScopedNames = workbook.Worksheets.Names.Filter(NameScopeType.Workbook, -1);

                // Output the names to the console
                Console.WriteLine($"Workbook‑scoped named ranges count: {workbookScopedNames.Length}");
                foreach (Name name in workbookScopedNames)
                {
                    // Name.Text contains the name identifier
                    Console.WriteLine(name.Text);
                }

                // (Optional) Save the workbook if you want to persist the changes
                // workbook.Save("WorkbookWithNames.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ListWorkbookScopedNames.Run();
        }
    }
}
