using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the workbook file
        string filePath = "input.xlsx";

        // Create LoadOptions and assign a custom LoadFilter that loads only the structure
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new StructureOnlyLoadFilter();

        // Load the workbook with the specified LoadOptions (no worksheet data will be loaded)
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Retrieve all defined names (both workbook‑scoped and worksheet‑scoped)
        Name[] definedNames = workbook.Worksheets.Names.Filter(NameScopeType.All, -1);

        // Display the names and their references
        foreach (Name name in definedNames)
        {
            Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}");
        }
    }

    // Custom LoadFilter that loads only the workbook structure for each sheet
    private class StructureOnlyLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load only the structure (no cell values) for every worksheet
            LoadDataFilterOptions = LoadDataFilterOptions.Structure;
        }
    }
}

// Author: Aspose.Cells .NET example