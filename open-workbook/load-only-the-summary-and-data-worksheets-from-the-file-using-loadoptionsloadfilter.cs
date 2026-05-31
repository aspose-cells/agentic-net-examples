using System;
using Aspose.Cells;

namespace AsposeCellsLoadSpecificSheets
{
    // Custom filter that loads data only for "Summary" and "Data" worksheets
    class CustomLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load full data for the required sheets, otherwise load only the structure
            if (sheet.Name == "Summary" || sheet.Name == "Data")
            {
                // Load everything (cells, formulas, formatting, etc.) for these sheets
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            }
            else
            {
                // Load only the sheet structure (no cell data) for all other sheets
                LoadDataFilterOptions = LoadDataFilterOptions.Structure;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourceFile = "input.xlsx";

            // Create LoadOptions and assign the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new CustomLoadFilter();

            // Load the workbook using the specified LoadOptions
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Demonstrate that only "Summary" and "Data" worksheets contain cell data
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet: {ws.Name}");
                Console.WriteLine($"  Cells count: {ws.Cells.Count}");
                // Show a sample cell value if data is present
                if (ws.Cells.Count > 0 && ws.Cells["A1"].Value != null)
                {
                    Console.WriteLine($"  A1 value: {ws.Cells["A1"].StringValue}");
                }
                else
                {
                    Console.WriteLine("  No cell data loaded.");
                }
            }

            // Save the workbook (optional, demonstrates the save rule)
            workbook.Save("output.xlsx");
        }
    }
}