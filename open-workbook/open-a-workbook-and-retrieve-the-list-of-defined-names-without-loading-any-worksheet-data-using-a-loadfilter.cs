// Title: Load an Excel workbook and list all defined names using Aspose.Cells LoadFilter (no worksheet data) – C# .NET
// Description: Shows how to create a LoadFilter with LoadDataFilterOptions.DefinedNames, apply it via LoadOptions, and open a workbook so that only named ranges are loaded. Retrieves both workbook‑ and worksheet‑scoped names using Worksheets.Names.Filter(NameScopeType.All, -1), prints each name with its RefersTo formula, and disposes the workbook.
// Keywords: Aspose.Cells LoadFilter | LoadDataFilterOptions.DefinedNames | C# load defined names only | retrieve named ranges .NET | open workbook without sheet data | Excel defined names extraction | memory‑efficient Excel loading | Worksheets.Names.Filter | NameScopeType.All
// Common Searches: Aspose.Cells load only defined names | C# read named ranges without loading worksheets | LoadFilter DefinedNames example | How to get all Excel named ranges using Aspose.Cells | Retrieve workbook scoped names Aspose.Cells .NET | List worksheet scoped names without cell data
// Developer Intent: The developer wants to open an Excel workbook and obtain every defined name while preventing any worksheet cell data from being loaded.
// Use Cases: Extract named ranges for metadata analysis without loading cell values, reducing memory usage. | Validate the presence of required named ranges in a template before processing large worksheets. | Generate documentation or migration reports of all defined names and their references without reading sheet content.
// AI Prompts: Write C# code with Aspose.Cells that loads only defined names from a workbook and returns them as a collection. | Show how to use LoadFilter with LoadDataFilterOptions.DefinedNames to list each defined name and its RefersTo string. | Explain how to modify the example to retrieve only workbook‑scoped names instead of all scopes.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadDefinedNames
{
    // Shows how to create a LoadFilter with LoadDataFilterOptions.DefinedNames, apply it via LoadOptions, and open a workbook so that only named ranges are loaded. Retrieves both workbook‑ and worksheet‑scoped names using Worksheets.Names.Filter(NameScopeType.All, -1), prints each name with its RefersTo formula, and disposes the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook file (replace with actual file path)
            string filePath = "input.xlsx";

            // Create a LoadFilter that loads only defined names (no worksheet data)
            LoadFilter nameOnlyFilter = new LoadFilter(LoadDataFilterOptions.DefinedNames);

            // Assign the filter to LoadOptions
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = nameOnlyFilter;

            // Load the workbook using the specified LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Retrieve all defined names (workbook and worksheet scope)
            Name[] allDefinedNames = workbook.Worksheets.Names.Filter(NameScopeType.All, -1);

            // Output the names and their references
            Console.WriteLine($"Total defined names loaded: {allDefinedNames.Length}");
            foreach (Name definedName in allDefinedNames)
            {
                Console.WriteLine($"Name: {definedName.Text}, RefersTo: {definedName.RefersTo}");
            }

            // Cleanup
            workbook.Dispose();
        }
    }
}
