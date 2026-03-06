using System;
using Aspose.Cells;

namespace AsposeCellsDefinedNamesFilter
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Path for the resulting workbook after loading without defined names
            string resultPath = "output.xlsx";

            // Create LoadOptions instance
            LoadOptions loadOptions = new LoadOptions();

            // Define filter options: load everything except defined names
            // LoadDataFilterOptions is a flag enum, so we combine All and then remove DefinedNames
            LoadDataFilterOptions filterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.DefinedNames;

            // Assign a LoadFilter with the custom options
            loadOptions.LoadFilter = new LoadFilter(filterOptions);

            // Load the workbook using the specified load options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // At this point, defined names are not loaded into the workbook
            // You can verify by checking the count of defined names
            Console.WriteLine("Defined names count after loading: " + workbook.Worksheets.Names.Count);

            // Save the workbook to a new file
            workbook.Save(resultPath);

            // Clean up
            workbook.Dispose();

            Console.WriteLine("Workbook saved without defined names to: " + resultPath);
        }
    }
}