using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = "input.xlsx";

        // Create a LoadFilter that loads everything except data validations
        // LoadDataFilterOptions.All includes all data; we remove DataValidation using bitwise NOT
        LoadDataFilterOptions filterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.DataValidation;
        LoadFilter loadFilter = new LoadFilter(filterOptions);

        // Assign the filter to LoadOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = loadFilter;

        // Load the workbook with the specified filter
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Demonstrate that worksheets are loaded and data validations are excluded
        Console.WriteLine("Number of worksheets loaded: " + workbook.Worksheets.Count);
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet '{sheet.Name}' - Cells: {sheet.Cells.Count}, Validations: {sheet.Validations.Count}");
        }

        // Save the workbook (optional, shows lifecycle usage)
        workbook.Save("output.xlsx");
    }
}