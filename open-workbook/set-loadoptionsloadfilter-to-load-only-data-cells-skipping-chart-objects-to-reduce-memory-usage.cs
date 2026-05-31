using System;
using Aspose.Cells;

namespace LoadOnlyCellDataExample
{
    class Program
    {
        static void Main()
        {
            // Create LoadOptions instance
            LoadOptions loadOptions = new LoadOptions();

            // Configure LoadFilter to load only cell data (values, formulas, formatting) and skip charts
            loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellData);

            // Load the workbook with the specified load options
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Example: verify that charts are not loaded
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet '{sheet.Name}' contains {sheet.Charts.Count} chart(s).");
            }

            // Save the workbook after loading (charts will be absent)
            workbook.Save("output.xlsx");
        }
    }
}