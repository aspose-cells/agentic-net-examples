using System;
using Aspose.Cells;

class LoadDefinedNamesDemo
{
    static void Main()
    {
        // Path to the source workbook file
        string sourcePath = "input.xlsx";

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Create a LoadFilter that loads only defined name objects
        LoadFilter filter = new LoadFilter(LoadDataFilterOptions.DefinedNames);
        loadOptions.LoadFilter = filter;

        // Load the workbook with the specified load options
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Retrieve all defined names (workbook and worksheet scope)
        Name[] definedNames = workbook.Worksheets.Names.Filter(NameScopeType.All, -1);
        Console.WriteLine("Defined names count: " + definedNames.Length);
        foreach (Name name in definedNames)
        {
            Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}");
        }

        // Save the workbook (optional verification)
        workbook.Save("output.xlsx");
    }
}