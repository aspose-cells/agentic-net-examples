using System;
using Aspose.Cells;

class LoadDefinedNamesExample
{
    static void Main()
    {
        // Path to the source Excel file
        string filePath = "input.xlsx";

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Configure the LoadFilter to load only defined names
        loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.DefinedNames);

        // Load the workbook with the specified load options
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Access the collection of defined names
        var definedNames = workbook.Worksheets.Names;

        // Output the number of defined names that were loaded
        Console.WriteLine("Number of defined names loaded: " + definedNames.Count);

        // Iterate through each defined name and display its details
        foreach (Name name in definedNames)
        {
            Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}");
        }

        // Save the workbook (optional, can be saved to a new file)
        workbook.Save("output.xlsx");
    }
}