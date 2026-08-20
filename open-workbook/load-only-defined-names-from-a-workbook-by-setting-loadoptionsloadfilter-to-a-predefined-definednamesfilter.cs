// Title: C# – Load Only Defined Names from an Excel Workbook Using Aspose.Cells LoadFilter
// Description: Demonstrates how to open an Excel file with Aspose.Cells while loading exclusively the workbook‑ and worksheet‑scoped defined names. The example configures LoadOptions.LoadFilter with LoadDataFilterOptions.DefinedNames, enumerates each name and its reference, and optionally saves the workbook.
// Keywords: Aspose.Cells LoadFilter DefinedNames | load defined names C# | retrieve named ranges Aspose.Cells | Excel defined names filter | LoadOptions LoadDataFilterOptions.DefinedNames example | enumerate named ranges C#
// Common Searches: Aspose.Cells load only defined names | C# load Excel named ranges without cell data | LoadFilter DefinedNames option tutorial | How to list all named ranges using Aspose.Cells | Open workbook with only name definitions in C#
// Developer Intent: Open an Excel workbook while loading only its defined names and iterate through them.
// Use Cases: Extract named ranges from a massive workbook without loading cell values to save memory. | Create an inventory of all workbook‑ and worksheet‑scoped names for documentation or validation. | Synchronize or copy defined names between workbooks after a lightweight load.
// AI Prompts: Show a C# code snippet that uses Aspose.Cells LoadOptions with LoadDataFilterOptions.DefinedNames to load only defined names and print each name with its RefersTo value. | Explain the performance impact of setting LoadFilter to DefinedNames when opening large Excel files with Aspose.Cells.

using Aspose.Cells;

// Demonstrates how to open an Excel file with Aspose.Cells while loading exclusively the workbook‑ and worksheet‑scoped defined names. The example configures LoadOptions.LoadFilter with LoadDataFilterOptions.DefinedNames, enumerates each name and its reference, and optionally saves the workbook.
class Program
{
    static void Main()
    {
        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Configure the LoadFilter to load only defined names
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.DefinedNames);

        // Load the workbook with the specified filter
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Retrieve all defined names (both workbook and worksheet scoped)
        Name[] definedNames = workbook.Worksheets.Names.Filter(NameScopeType.All, -1);

        // Output each defined name and its reference
        foreach (Name name in definedNames)
        {
            System.Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}");
        }

        // Save the workbook (optional, to verify that loading succeeded)
        workbook.Save("output.xlsx");
    }
}
