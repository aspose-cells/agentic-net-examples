// Title: Aspose.Cells C# – Load Only Defined Names Using LoadFilter (DefinedNames)
// Description: Demonstrates how to use Aspose.Cells LoadOptions with LoadFilter set to LoadDataFilterOptions.DefinedNames to open an Excel workbook while loading only its named ranges. The example enumerates both workbook‑ and worksheet‑scoped names, prints their references, and shows an optional save operation. This approach reduces memory usage when only name metadata is required.
// Keywords: Aspose.Cells | LoadOptions | LoadFilter | DefinedNames | named ranges | C# | .NET | Excel workbook | retrieve defined names | filter Excel load | performance optimization
// Common Searches: Aspose.Cells load only defined names C# | LoadFilter DefinedNames example | How to read named ranges without cell data using Aspose.Cells | Get all workbook scoped names Aspose.Cells .NET | Filter Excel load to named ranges Aspose
// Developer Intent: Load an Excel file with Aspose.Cells so that only its defined names are read, then list those names.
// Use Cases: List or validate named ranges without loading cell contents to improve memory and speed. | Extract metadata for documentation or auditing of workbook naming conventions. | Transfer or copy defined names to another workbook after a lightweight load.
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells using LoadOptions.LoadFilter = LoadDataFilterOptions.DefinedNames and prints each defined name and its reference. | Show how to retrieve only workbook‑scoped defined names after loading a workbook with the DefinedNames filter in Aspose.Cells. | Explain how to save a workbook after loading only defined names, ensuring that no cell data is written to the output file.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells LoadOptions with LoadFilter set to LoadDataFilterOptions.DefinedNames to open an Excel workbook while loading only its named ranges. The example enumerates both workbook‑ and worksheet‑scoped names, prints their references, and shows an optional save operation. This approach reduces memory usage when only name metadata is required.
class LoadDefinedNamesOnly
{
    static void Main()
    {
        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Set LoadFilter to load only defined names
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.DefinedNames);

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("InputWorkbook.xlsx", loadOptions);

        // Retrieve all defined names (both workbook and worksheet scoped)
        Name[] allDefinedNames = workbook.Worksheets.Names.Filter(NameScopeType.All, -1);

        // Example: output the names to console
        foreach (Name name in allDefinedNames)
        {
            Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}");
        }

        // Save the workbook (optional, to verify loading)
        workbook.Save("OutputWorkbook.xlsx");
    }
}
