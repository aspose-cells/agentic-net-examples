using System;
using Aspose.Cells;

class FilterDefinedNamesDemo
{
    static void Main()
    {
        // Path to the source workbook (replace with your actual file)
        string sourcePath = "Source.xlsx";

        // Path where the processed workbook will be saved
        string outputPath = "FilteredNames.xlsx";

        // Create LoadOptions and assign a custom LoadFilter that loads only defined names
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new DefinedNamesOnlyLoadFilter();

        // Load the workbook using the specified LoadOptions
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Retrieve all defined names (both workbook‑scoped and worksheet‑scoped)
        Name[] allNames = workbook.Worksheets.Names.Filter(NameScopeType.All, -1);

        // Output information about the loaded names
        Console.WriteLine($"Loaded {allNames.Length} defined name(s):");
        foreach (Name name in allNames)
        {
            Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}");
        }

        // Save the workbook (structure and names are preserved)
        workbook.Save(outputPath);
    }

    // Custom LoadFilter that restricts loading to defined name objects only
    private class DefinedNamesOnlyLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load only the defined names; other data (cells, formats, etc.) is ignored
            LoadDataFilterOptions = LoadDataFilterOptions.DefinedNames;
        }
    }
}