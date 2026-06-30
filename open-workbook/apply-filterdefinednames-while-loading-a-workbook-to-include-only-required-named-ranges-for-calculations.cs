using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – load workbook with filtered named ranges
class Program
{
    static void Main()
    {
        // Define the named ranges that are required for calculations
        string[] requiredNames = { "TotalSales", "TaxRate" };

        // Configure load options with a custom LoadFilter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new RequiredNamesLoadFilter(requiredNames);

        // Load the workbook using the specified options
        Workbook workbook = new Workbook("Input.xlsx", loadOptions);

        // Retrieve all workbook‑scoped names
        Name[] allNames = workbook.Worksheets.Names.Filter(NameScopeType.Workbook, -1);

        // Process only the required named ranges
        foreach (Name name in allNames)
        {
            if (Array.IndexOf(requiredNames, name.Text) >= 0)
            {
                Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}");
                // Additional calculation logic can be placed here
            }
        }

        // Save the workbook (if any modifications were made)
        workbook.Save("Output.xlsx");
    }

    // Custom LoadFilter that limits loading to structure only (names are part of the structure)
    private class RequiredNamesLoadFilter : LoadFilter
    {
        private readonly string[] _requiredNames;

        public RequiredNamesLoadFilter(string[] requiredNames)
        {
            _requiredNames = requiredNames;
        }

        public override void StartSheet(Worksheet sheet)
        {
            // Load only the workbook structure for each sheet; data cells are not needed for name filtering
            LoadDataFilterOptions = LoadDataFilterOptions.Structure;
        }
    }
}