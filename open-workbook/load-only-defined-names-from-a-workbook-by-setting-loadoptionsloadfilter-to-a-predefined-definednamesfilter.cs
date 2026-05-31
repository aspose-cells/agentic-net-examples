// Load only defined names from a workbook using a predefined filter
using Aspose.Cells;
using System;

class LoadDefinedNamesOnly
{
    static void Main()
    {
        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Set LoadFilter to load only defined names
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.DefinedNames);

        // Load the workbook with the specified filter
        Workbook workbook = new Workbook("InputWorkbook.xlsx", loadOptions);

        // Example: iterate through all defined names (workbook and worksheet scope)
        foreach (Name name in workbook.Worksheets.Names)
        {
            Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}");
        }

        // Save the workbook (the loaded data includes only defined names)
        workbook.Save("OutputWorkbook.xlsx");
    }
}