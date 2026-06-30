using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example
    class LoadWorkbookExcludingTempSheets
    {
        static void Main()
        {
            // Create LoadOptions and assign a custom LoadFilter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new TempExcludingLoadFilter();

            // Load the workbook with the specified options
            // The file path should point to an existing Excel file
            string sourcePath = "SourceWorkbook.xlsx";
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the loaded workbook (optional, to verify the result)
            workbook.Save("FilteredWorkbook.xlsx");
        }

        // Custom LoadFilter that skips loading data for worksheets named "Temp"
        private class TempExcludingLoadFilter : LoadFilter
        {
            public override void StartSheet(Worksheet sheet)
            {
                if (sheet.Name.Equals("Temp", StringComparison.OrdinalIgnoreCase))
                {
                    // Load only the sheet structure (no cell data) for "Temp" sheets
                    LoadDataFilterOptions = LoadDataFilterOptions.Structure;
                }
                else
                {
                    // Load full data for all other sheets
                    LoadDataFilterOptions = LoadDataFilterOptions.All;
                }
            }

            // NOTE: If a dedicated ExcludedSheets collection existed on LoadFilter,
            // you would add "Temp" to that collection here.
            // Since the current API documentation does not expose such a member,
            // the above approach uses LoadDataFilterOptions to effectively ignore data
            // in sheets named "Temp". Replace with ExcludedSheets usage when available.
        }
    }
}