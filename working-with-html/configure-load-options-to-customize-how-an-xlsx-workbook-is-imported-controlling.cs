using System;
using System.Globalization;
using Aspose.Cells;

namespace LoadOptionsDemo
{
    class Program
    {
        static void Main()
        {
            // Create LoadOptions instance using the default constructor
            LoadOptions loadOptions = new LoadOptions();

            // Configure parsing behavior
            loadOptions.ParsingFormulaOnOpen = true;          // Parse formulas while loading
            loadOptions.ParsingPivotCachedRecords = false;   // Do not parse pivot cached records

            // Configure data handling
            loadOptions.KeepUnparsedData = true;             // Keep unparsed data in memory
            loadOptions.MemorySetting = MemorySetting.MemoryPreference; // Optimize memory usage

            // Set a password if the workbook is protected
            loadOptions.Password = "securePassword";

            // Apply a load filter to load only cell data and charts
            loadOptions.LoadFilter = new LoadFilter(
                LoadDataFilterOptions.CellData | LoadDataFilterOptions.Chart);

            // Optionally set auto fitter options (e.g., auto‑fit merged cells)
            loadOptions.AutoFitterOptions = new AutoFitterOptions { AutoFitMergedCells = true };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Perform a simple operation to demonstrate that the workbook is loaded
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Loaded with custom LoadOptions");

            // Save the workbook to a new file
            workbook.Save("output.xlsx");
        }
    }
}