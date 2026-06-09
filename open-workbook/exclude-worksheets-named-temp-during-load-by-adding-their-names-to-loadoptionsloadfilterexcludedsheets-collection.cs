using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Custom LoadFilter that excludes worksheets named "Temp"
    class TempExcludingLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // If the worksheet name is "Temp", skip loading any data for it
            if (string.Equals(sheet.Name, "Temp", StringComparison.OrdinalIgnoreCase))
            {
                // Do not load any data for this sheet
                LoadDataFilterOptions = LoadDataFilterOptions.None;
            }
            else
            {
                // Load everything for other sheets
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Prepare load options and assign the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new TempExcludingLoadFilter();

            // Load the workbook using the specified options
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Save the workbook; the "Temp" sheets will be excluded from loading
            workbook.Save("output.xlsx");
        }
    }
}