// Title: C# – Load only cell values with Aspose.Cells LoadFilter, omit charts for lower memory
// Description: The sample creates a LoadOptions object, assigns a LoadFilter using LoadDataFilterOptions.CellData, disables KeepUnparsedData, opens the source workbook without any chart objects, and saves a new file. This approach trims the in‑memory footprint when processing Excel files.
// Keywords: Aspose.Cells LoadFilter | LoadDataFilterOptions.CellData | skip charts | memory optimization | C# workbook loading | Excel without charts | reduce RAM usage | LoadOptions KeepUnparsedData false | lightweight Excel copy | process large workbooks
// Common Searches: Aspose.Cells load workbook without charts | C# load only cell data from Excel | how to lower memory when opening large Excel files with Aspose | LoadFilter CellData example | skip chart objects Aspose.Cells | effect of KeepUnparsedData false | create chart‑free copy of Excel using Aspose | minimal memory usage LoadOptions
// Developer Intent: Open an Excel file and retrieve just the cell contents, excluding charts, to keep RAM consumption low.
// Use Cases: A data‑processing service reads thousands of spreadsheets that contain many charts but only needs the numeric values; using CellData filtering avoids loading unnecessary graphics. | An ETL pipeline extracts cell formulas from large workbooks while discarding visual elements to stay within server memory limits. | A reporting tool generates a lightweight export of a workbook for downstream systems, removing all chart objects before saving.
// AI Prompts: Generate C# code that uses Aspose.Cells LoadOptions with LoadFilter to load only cell data and save the workbook without charts. | Explain which workbook components are omitted when LoadDataFilterOptions.CellData is applied and how this impacts memory usage. | Show how combining LoadFilter with KeepUnparsedData = false further reduces the memory footprint during workbook loading.

using System;
using Aspose.Cells;

namespace LoadOnlyCellDataExample
{
    // The sample creates a LoadOptions object, assigns a LoadFilter using LoadDataFilterOptions.CellData, disables KeepUnparsedData, opens the source workbook without any chart objects, and saves a new file. This approach trims the in‑memory footprint when processing Excel files.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook that may contain charts
            string sourcePath = "input.xlsx";

            // Path where the filtered workbook will be saved
            string destinationPath = "output.xlsx";

            // Create LoadOptions instance
            LoadOptions loadOptions = new LoadOptions();

            // Configure LoadFilter to load only cell data (values, formulas, formatting) and skip charts
            // LoadDataFilterOptions.CellData includes cell values, formulas and formatting but excludes Chart objects
            loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellData);

            // Optional: reduce memory usage further by not keeping unparsed data
            loadOptions.KeepUnparsedData = false;

            // Load the workbook with the specified load options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Perform any additional processing here if needed
            // For demonstration, output the number of worksheets loaded
            Console.WriteLine($"Worksheets loaded: {workbook.Worksheets.Count}");

            // Save the workbook; charts will not be present in the saved file
            workbook.Save(destinationPath);
        }
    }
}
