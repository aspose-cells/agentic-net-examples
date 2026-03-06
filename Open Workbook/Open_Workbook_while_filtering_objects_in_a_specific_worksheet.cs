using System;
using Aspose.Cells;

namespace AsposeCellsLoadFilterDemo
{
    // Custom load filter to control loading of each worksheet
    class CustomLoadFilter : LoadFilter
    {
        private readonly string _targetSheetName;

        public CustomLoadFilter(string targetSheetName)
        {
            _targetSheetName = targetSheetName;
        }

        public override void StartSheet(Worksheet sheet)
        {
            if (sheet != null && sheet.Name.Equals(_targetSheetName, StringComparison.OrdinalIgnoreCase))
            {
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            }
            else
            {
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourceFile = "SourceWorkbook.xlsx";

            // Create a LoadOptions instance and assign the custom filter
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFilter = new CustomLoadFilter("TargetSheet")
            };

            // Load the workbook using the LoadOptions (the filter will be applied)
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Get the target sheet (if it exists)
            Worksheet targetSheet = workbook.Worksheets["TargetSheet"];
            if (targetSheet != null)
            {
                Console.WriteLine($"Target sheet '{targetSheet.Name}' cell count: {targetSheet.Cells.Count}");
            }
            else
            {
                Console.WriteLine("Target sheet 'TargetSheet' not found.");
            }

            // Demonstrate that other sheets have only structure loaded (no cell data)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (!string.Equals(sheet.Name, "TargetSheet", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"Other sheet '{sheet.Name}' cell count (should be 0): {sheet.Cells.Count}");
                }
            }

            // Save the workbook (the data that was loaded remains unchanged)
            workbook.Save("FilteredWorkbook.xlsx");
        }
    }
}