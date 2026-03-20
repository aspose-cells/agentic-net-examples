using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source workbook file
        string sourcePath = "input.xlsx";

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Define filter options that load everything except defined names
        // LoadDataFilterOptions.All includes all data; we remove the DefinedNames flag using bitwise NOT
        LoadDataFilterOptions filterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.DefinedNames;

        // Create a LoadFilter with the custom filter options
        LoadFilter loadFilter = new LoadFilter(filterOptions);

        // Assign the filter to the LoadOptions
        loadOptions.LoadFilter = loadFilter;

        // Load the workbook using the file path and the configured LoadOptions
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Verify that no defined names were loaded (should be zero)
        Name[] definedNames = workbook.Worksheets.Names.Filter(NameScopeType.All, -1);
        Console.WriteLine("Number of defined names loaded: " + definedNames.Length);

        // Save the workbook to a new file (optional verification)
        workbook.Save("output.xlsx");
    }
}