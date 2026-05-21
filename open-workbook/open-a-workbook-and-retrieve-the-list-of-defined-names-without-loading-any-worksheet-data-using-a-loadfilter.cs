using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class RetrieveDefinedNames
    {
        static void Main()
        {
            // Path to the source workbook
            string sourceFile = "input.xlsx";

            // Create a LoadFilter that loads only defined names (no worksheet data)
            LoadFilter nameOnlyFilter = new LoadFilter(LoadDataFilterOptions.DefinedNames);

            // Assign the filter to LoadOptions
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = nameOnlyFilter;

            // Load the workbook using the specified LoadOptions
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Retrieve all defined names (workbook and worksheet scope)
            Name[] allNames = workbook.Worksheets.Names.Filter(NameScopeType.All, -1);

            // Output the names and their references
            Console.WriteLine($"Total defined names: {allNames.Length}");
            foreach (Name name in allNames)
            {
                Console.WriteLine($"Name: {name.Text}, RefersTo: {name.RefersTo}");
            }

            // Cleanup
            workbook.Dispose();
        }
    }
}